/* eslint-disable max-lines-per-function, complexity */
import React, { useEffect, useState } from 'react';
import { useAuth as useAuthCtx } from '../../context/AuthContext';
import { useFeatureFlags } from '../../context/FeatureFlagsContext';
import Alerts, { alertSuccess, alertError, alertInfo, alertWarning, showToast } from '../../utils/alerts';
import { getApiBase, isSQLiteEnabled, isAdminLight, useAdminModalSelectors as adminModalSelectorsDefault } from '../../utils/runtimeConfig';
import { getGraphToken } from '../../services/authTokens';

type AdminSettingsProps = { canEdit: boolean };

const AdminSettings: React.FC<AdminSettingsProps> = ({ canEdit }) => {
  const { account } = useAuthCtx();
  const { refresh: refreshFlags } = useFeatureFlags();
  const storageKey = 'admin_settings';
  const [settings, setSettings] = useState({
    enableUpload: false,
    requireSig: false,
    autoReminder: true,
    reminderDays: 3,
    allowBulkAssignment: true,
    requireApproval: false
  });

  // External support flag (server-backed)
  const [extEnabled, setExtEnabled] = useState<boolean>(false);
  const [extLoading, setExtLoading] = useState<boolean>(false);
  const [extSaving, setExtSaving] = useState<boolean>(false);
  const apiBase = (getApiBase() as string) || '';

  // Legal consent document
  const [legalDoc, setLegalDoc] = useState<{ fileId: number | null; url: string | null; name: string | null }>({ fileId: null, url: null, name: null });
  const [legalBusy, setLegalBusy] = useState<boolean>(false);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) {
        const obj = JSON.parse(raw);
        setSettings(prev => ({ ...prev, ...obj }));
      }
  } catch { /* ignore */ }
  }, []);

  // Load external support flag
  useEffect(() => {
    (async () => {
      try {
        setExtLoading(true);
        if (!apiBase) { setExtEnabled(false); return; }
        const res = await fetch(`${apiBase}/api/settings/external-support`, { cache: 'no-store' });
        const j = await res.json();
        setExtEnabled(!!j?.enabled);
      } catch {
        setExtEnabled(false);
      } finally {
        setExtLoading(false);
      }
    })();
  }, [apiBase]);

  // Load current legal consent document
  useEffect(() => {
    (async () => {
      try {
        if (!apiBase) return;
        const res = await fetch(`${apiBase}/api/settings/legal-consent`, { cache: 'no-store' });
        const j = await res.json();
        setLegalDoc({ fileId: j?.fileId ?? null, url: j?.url ? (apiBase + j.url) : null, name: j?.name ?? null });
  } catch { /* ignore */ }
    })();
  }, [apiBase]);

  const apply = () => {
    if (!canEdit) return;
    try {
      localStorage.setItem(storageKey, JSON.stringify(settings));
      Alerts.toast('Settings saved');
    } catch (e) {
      console.warn(e);
    }
  };

  const saveExternalSupport = async (value: boolean) => {
    if (!canEdit || !apiBase) return;
    setExtSaving(true);
    try {
      const res = await fetch(`${apiBase}/api/settings/external-support`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ enabled: value }) });
      if (!res.ok) throw new Error('save_failed');
      setExtEnabled(value);
  try { await refreshFlags(); } catch { /* ignore */ }
      Alerts.toast(`External user support ${value ? 'enabled' : 'disabled'}`);
    } catch {
      Alerts.toast('Failed to save external support setting');
    } finally {
      setExtSaving(false);
    }
  };

  const seedSqliteForMe = async () => {
    try {
      if (!isSQLiteEnabled()) {
        alertWarning('SQLite disabled', 'Enable SQLite (REACT_APP_ENABLE_SQLITE=true) and set REACT_APP_API_BASE to seed.');
        return;
      }
      if (!account?.username) {
        alertInfo('Sign in required', 'Sign in first to seed data for your account.');
        return;
      }
      const base = (getApiBase() as string);
      const res = await fetch(`${base}/api/seed?email=${encodeURIComponent(account.username)}`, { method: 'POST' });
      if (!res.ok) throw new Error('Seed failed');
      const j = await res.json().catch(() => ({}));
      alertSuccess('SQLite seeded', `BatchId: <b>${j?.batchId ?? 'n/a'}</b>`);
    } catch (e) {
      alertError('Seed failed', 'Unable to seed demo data.');
    }
  };

  const grantCorePermissions = async () => {
    try {
      // Request common scopes used across the Admin feature set
      // Do these in series to present clearer consent prompts
      await getGraphToken(['User.Read']);
      await getGraphToken(['Group.Read.All']);
      await getGraphToken(['Sites.Read.All','Files.ReadWrite.All']);
      await getGraphToken(['Mail.Send']);
      showToast('Core Graph permissions granted');
    } catch (e) {
      showToast('Grant permissions failed', 'error');
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <h3 style={{ margin: 0, fontSize: 16 }}>System Settings</h3>
      {/* External Support Toggle */}
      <div className="card" style={{ padding: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontWeight: 700 }}>External User Support</div>
            <div className="small muted">When disabled, external login, onboarding, and related UI/routes are hidden.</div>
          </div>
          <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <input type="checkbox" checked={!!extEnabled} disabled={extLoading || extSaving || !canEdit} onChange={e => saveExternalSupport(e.target.checked)} />
            <span>{extEnabled ? 'Enabled' : 'Disabled'}</span>
          </label>
        </div>
      </div>
      {/* Legal Consent Document */}
      <div className="card" style={{ padding: 16 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
          <div>
            <div style={{ fontWeight: 700 }}>Legal Consent Document</div>
            <div className="small muted">PDF shown to users from the consent dialog. Applies globally.</div>
            {legalDoc?.url ? (
              <div className="small" style={{ marginTop: 6 }}>
                Current: <a href={legalDoc.url} target="_blank" rel="noreferrer">{legalDoc.name || 'document.pdf'} ↗</a>
              </div>
            ) : (
              <div className="small muted" style={{ marginTop: 6 }}>Not set</div>
            )}
          </div>
          <div>
            <label className="btn sm" style={{ cursor: canEdit ? 'pointer' : 'not-allowed', opacity: canEdit ? 1 : .6 }}>
              {legalBusy ? 'Uploading…' : (legalDoc?.fileId ? 'Replace PDF' : 'Upload PDF')}
              <input type="file" accept="application/pdf" style={{ display: 'none' }} disabled={!canEdit || legalBusy} onChange={async (e) => {
                try {
                  const file = e.target.files && e.target.files[0];
                  if (!file || !apiBase) return;
                  setLegalBusy(true);
                  const fd = new FormData();
                  fd.append('file', file);
                  const up = await fetch(`${apiBase}/api/files/upload`, { method: 'POST', body: fd });
                  const uj = await up.json();
                  if (!up.ok || !uj?.id) { showToast('Upload failed', 'error'); return; }
                  const put = await fetch(`${apiBase}/api/settings/legal-consent`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fileId: uj.id }) });
                  if (!put.ok) { showToast('Save failed', 'error'); return; }
                  setLegalDoc({ fileId: uj.id, url: `${apiBase}/api/files/${uj.id}`, name: file.name });
                  showToast('Legal document saved', 'success');
                } catch {
                  showToast('Upload failed', 'error');
                } finally {
                  setLegalBusy(false);
                  try { (e.target as HTMLInputElement).value = ''; } catch { /* ignore */ }
                }
              }} />
            </label>
            {legalDoc?.fileId && (
              <button className="btn ghost sm" style={{ marginLeft: 8 }} disabled={!canEdit || legalBusy} onClick={async () => {
                try {
                  if (!apiBase) return;
                  setLegalBusy(true);
                  const put = await fetch(`${apiBase}/api/settings/legal-consent`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fileId: null }) });
                  if (!put.ok) { showToast('Failed to clear', 'error'); return; }
                  setLegalDoc({ fileId: null, url: null, name: null });
                  showToast('Cleared legal document', 'success');
                } catch { showToast('Failed to clear', 'error'); }
                finally { setLegalBusy(false); }
              }}>Clear</button>
            )}
          </div>
        </div>
      </div>

      <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.enableUpload} onChange={e => setSettings({...settings, enableUpload: e.target.checked})} disabled={!canEdit} />
          <span className="small">Enable document upload</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.requireSig} onChange={e => setSettings({...settings, requireSig: e.target.checked})} disabled={!canEdit} />
          <span className="small">Require digital signatures</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.autoReminder} onChange={e => setSettings({...settings, autoReminder: e.target.checked})} disabled={!canEdit} />
          <span className="small">Auto-send reminders</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.allowBulkAssignment} onChange={e => setSettings({...settings, allowBulkAssignment: e.target.checked})} disabled={!canEdit} />
          <span className="small">Allow bulk assignments</span>
        </label>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <span className="small">Reminder frequency:</span>
        <select value={settings.reminderDays} onChange={e => setSettings({...settings, reminderDays: parseInt(e.target.value)})} disabled={!canEdit}>
          <option value={1}>Daily</option>
          <option value={3}>Every 3 days</option>
          <option value={7}>Weekly</option>
          <option value={14}>Bi-weekly</option>
        </select>
      </div>

      <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        {canEdit && <button className="btn" onClick={apply}>Save Settings</button>}
        {!canEdit && <span className="small muted">Read-only access</span>}
        {canEdit && <button className="btn ghost" onClick={seedSqliteForMe} title="Seed SQLite with a demo batch, docs, and recipients for your account">Seed SQLite (for me)</button>}
        {canEdit && <button className="btn ghost" onClick={grantCorePermissions} title="Request common Microsoft Graph permissions in one go">Grant Core Permissions</button>}
      </div>

      {/* Environment summary */}
      <div className="card" style={{ padding: 12 }}>
        <div style={{ fontWeight: 700, marginBottom: 6 }}>Environment & Feature Flags</div>
        <div className="small" style={{ display: 'grid', gridTemplateColumns: '1fr 2fr', rowGap: 4, columnGap: 8 }}>
          <div>SQLite Enabled</div><div>{isSQLiteEnabled() ? 'true' : 'false'}</div>
          <div>API Base</div><div>{String(getApiBase() || '—')}</div>
          <div>Admin Light Mode</div><div>{isAdminLight() ? 'true' : 'false'}</div>
          <div>Modal Selectors (default)</div><div>{adminModalSelectorsDefault() ? 'true' : 'false'}</div>
        </div>
      </div>
    </div>
  );
};

export default AdminSettings;
