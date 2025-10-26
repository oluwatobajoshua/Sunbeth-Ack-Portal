/* eslint-disable max-lines-per-function, complexity */
import React, { useEffect, useMemo, useState } from 'react';
import { getApiBase } from '../../utils/runtimeConfig';
import { showToast } from '../../utils/alerts';
import { useAuth } from '../../context/AuthContext';

// Helper to extract a concise backend error message without deep nesting
const parseErrorMessage = async (res: Response, fallback: string) => {
  try {
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('application/json')) {
      const e = await res.json().catch(() => null);
      if (e && typeof e === 'object' && 'error' in (e as any) && (e as any).error) {
        return `${fallback} (${(e as any).error})`;
      }
      return fallback;
    }
    const t = await res.text().catch(() => '');
    return t ? `${fallback} (${t.substring(0, 140)})` : fallback;
  } catch {
    return fallback;
  }
};

export type Policy = {
  id: number;
  name: string;
  description?: string | null;
  frequency: 'daily' | 'weekly' | 'monthly' | 'quarterly' | 'semiannual' | 'annual' | 'custom';
  intervalDays?: number | null;
  required: boolean;
  fileId: number;
  fileIds?: number[];
  sha256?: string | null;
  active: boolean;
  startOn?: string | null;
  dueInDays?: number | null;
  graceDays?: number | null;
  createdAt?: string;
  updatedAt?: string;
};

const freqOptions = [
  'daily','weekly','monthly','quarterly','semiannual','annual','custom'
] as const;

const Policies: React.FC = () => {
  const apiBase = (getApiBase() as string) || '';
  const { account } = useAuth();
  const adminEmail = (account?.username || '').trim();
  const [policies, setPolicies] = useState<Policy[]>([]);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);

  const [q, setQ] = useState('');
  const [files, setFiles] = useState<Array<{ id: number; name: string; url: string; size?: number; sha256?: string }>>([]);
  const [form, setForm] = useState<Partial<Policy> & { fileIds?: number[] }>({ frequency: 'annual', required: true, active: true, dueInDays: 30, graceDays: 0, fileIds: [] });

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return policies;
    return policies.filter(p => p.name.toLowerCase().includes(s));
  }, [q, policies]);

  // Only attach custom admin header when it won't trigger CORS preflight across origins
  const adminHeaders = useMemo(() => {
    if (!adminEmail) return {} as Record<string, string>;
    try {
      const base = apiBase || '';
      const isAbs = /^https?:\/\//i.test(base);
      if (!isAbs) return { 'X-Admin-Email': adminEmail };
      const same = typeof window !== 'undefined' && base.startsWith(window.location.origin);
      return same ? { 'X-Admin-Email': adminEmail } : {};
    } catch {
      return {} as Record<string, string>;
    }
  }, [apiBase, adminEmail]);

  const load = async () => {
    setLoading(true);
    try {
      const url = `${apiBase}/api/admin/policies${adminEmail ? `?adminEmail=${encodeURIComponent(adminEmail)}` : ''}`;
      const res = await fetch(url, { cache: 'no-store', headers: Object.keys(adminHeaders).length ? adminHeaders : undefined });
      const j = await res.json();
      setPolicies(Array.isArray(j?.policies) ? j.policies : []);
    } catch { setPolicies([]); }
    finally { setLoading(false); }
  };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => { if (apiBase) load(); }, [apiBase, adminEmail]);

  const loadFiles = async () => {
    try {
      const res = await fetch(`${apiBase}/api/library/list?limit=200`, { cache: 'no-store' });
      const j = await res.json();
      const arr = Array.isArray(j?.files) ? j.files : [];
      setFiles(arr.map((r: any) => ({ id: Number(r.id), name: String(r.name || 'file'), url: `${apiBase}${r.url}`, size: Number(r.size)||undefined, sha256: r.sha256 })));
    } catch { setFiles([]); }
  };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => { if (apiBase) loadFiles(); }, [apiBase]);

  const save = async () => {
    const selected = Array.isArray(form.fileIds) ? form.fileIds : (form.fileId ? [form.fileId] : []);
    if (!form.name || selected.length === 0) { showToast('Name and at least one file are required', 'warning'); return; }
    setSaving(true);
    try {
      const url = `${apiBase}/api/admin/policies${adminEmail ? `?adminEmail=${encodeURIComponent(adminEmail)}` : ''}`;
      const res = await fetch(url, {
        method: 'POST', headers: { 'Content-Type': 'application/json', ...adminHeaders },
        body: JSON.stringify({
          name: form.name,
          description: form.description || null,
          frequency: form.frequency || 'annual',
          intervalDays: form.frequency === 'custom' ? (form.intervalDays || 0) : null,
          required: !!form.required,
          fileIds: selected,
          startOn: form.startOn || null,
          dueInDays: form.dueInDays || 30,
          graceDays: form.graceDays || 0,
          active: form.active !== false
        })
      });
  if (!res.ok) throw new Error(await parseErrorMessage(res, 'Failed to save policy'));
      setForm({ frequency: 'annual', required: true, active: true, dueInDays: 30, graceDays: 0, fileIds: [] });
      await load();
      showToast('Policy saved', 'success');
    } catch (e: any) { showToast(String(e?.message || 'Failed to save policy'), 'error'); }
    finally { setSaving(false); }
  };

  const del = async (id: number) => {
    try {
      const url = `${apiBase}/api/admin/policies/${id}${adminEmail ? `?adminEmail=${encodeURIComponent(adminEmail)}` : ''}`;
      const res = await fetch(url, { method: 'DELETE', headers: Object.keys(adminHeaders).length ? adminHeaders : undefined });
      if (!res.ok) throw new Error('delete_failed');
      await load();
      showToast('Policy deleted', 'success');
    } catch { showToast('Delete failed', 'error'); }
  };

  return (
    <div className="card" style={{ padding: 16 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <h3 style={{ margin: 0, fontSize: 16 }}>Policies (Recurring Acknowledgements)</h3>
          <div className="small muted">Define mandatory annual (or other) acknowledgements per document; applies to all employees by default.</div>
        </div>
        <input placeholder="Search policies..." value={q} onChange={e => setQ(e.target.value)} />
      </div>

      {/* Create */}
      <div className="card" style={{ marginTop: 12, padding: 12 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <div>
            <label className="small" htmlFor="policyName">Policy Name</label>
            <input id="policyName" value={form.name || ''} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="Code of Conduct (Annual)" />
          </div>
          <div>
            <label className="small" htmlFor="policyFreq">Frequency</label>
            <select id="policyFreq" value={String(form.frequency || 'annual')} onChange={e => setForm({ ...form, frequency: e.target.value as any })}>
              {freqOptions.map(f => <option key={f} value={f}>{f}</option>)}
            </select>
          </div>
          {form.frequency === 'custom' && (
            <div>
              <label className="small" htmlFor="intervalDays">Interval (days)</label>
              <input id="intervalDays" type="number" value={Number(form.intervalDays||0)} onChange={e => setForm({ ...form, intervalDays: Number(e.target.value)||0 })} />
            </div>
          )}
          <div>
            <div className="small">Required</div>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <input type="checkbox" checked={form.required !== false} onChange={e => setForm({ ...form, required: e.target.checked })} /> Required for all employees
            </label>
          </div>
          <div>
            <label className="small" htmlFor="startOn">Start On (optional)</label>
            <input id="startOn" type="date" value={form.startOn || ''} onChange={e => setForm({ ...form, startOn: e.target.value })} />
          </div>
          <div>
            <label className="small" htmlFor="dueDays">Due Window (days)</label>
            <input id="dueDays" type="number" value={Number(form.dueInDays||30)} onChange={e => setForm({ ...form, dueInDays: Number(e.target.value)||0 })} />
          </div>
          <div>
            <label className="small" htmlFor="graceDays">Grace (days)</label>
            <input id="graceDays" type="number" value={Number(form.graceDays||0)} onChange={e => setForm({ ...form, graceDays: Number(e.target.value)||0 })} />
          </div>
        </div>
        <div style={{ marginTop: 12 }}>
          <div className="small muted" style={{ marginBottom: 6 }}>Select document(s) from Library</div>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, maxHeight: 180, overflowY: 'auto' }}>
            {files.map(f => {
              const selected = Array.isArray(form.fileIds) ? form.fileIds.includes(f.id) : form.fileId === f.id;
              return (
                <label key={f.id} className="small" style={{ display: 'grid', gridTemplateColumns: 'auto 1fr auto', gap: 8, alignItems: 'center', padding: '6px 8px', border: '1px solid #f0f0f0', borderRadius: 6 }}>
                  <input type="checkbox" checked={!!selected} onChange={e => {
                    const cur = new Set<number>(Array.isArray(form.fileIds) ? form.fileIds : (form.fileId ? [form.fileId] : []));
                    if (e.target.checked) cur.add(f.id); else cur.delete(f.id);
                    setForm({ ...form, fileIds: Array.from(cur) });
                  }} />
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                    <div className="muted" style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                      {f.size ? <span>{(f.size/1024/1024).toFixed(1)} MB</span> : null}
                      <a href={f.url} target="_blank" rel="noreferrer">Preview ↗</a>
                    </div>
                  </div>
                  <span className="badge">file #{f.id}</span>
                </label>
              );
            })}
          </div>
          <div className="small" style={{ marginTop: 6 }}>
            Selected: <strong>{Array.isArray(form.fileIds) ? form.fileIds.length : (form.fileId ? 1 : 0)}</strong>
          </div>
        </div>
        <div style={{ marginTop: 12, textAlign: 'right' }}>
          <button className="btn" onClick={save} disabled={saving}>Save Policy</button>
        </div>
      </div>

      {/* List */}
      <div className="card" style={{ marginTop: 12, padding: 12 }}>
        <div style={{ fontWeight: 700, marginBottom: 6 }}>Existing Policies</div>
        {loading ? <div className="small muted">Loading...</div> : (
          filtered.length === 0 ? <div className="small muted">No policies found.</div> : (
            <div style={{ display: 'grid', gap: 8 }}>
              {filtered.map(p => (
                <div key={p.id} style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8, alignItems: 'center', border: '1px solid #f0f0f0', borderRadius: 6, padding: 8 }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{p.name}</div>
                    <div className="small muted" style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                      <span>freq: {p.frequency}{p.frequency==='custom' && p.intervalDays ? ` (${p.intervalDays} days)` : ''}</span>
                      <span>required: {p.required ? 'yes' : 'no'}</span>
                      <span>active: {p.active ? 'yes' : 'no'}</span>
                      {p.startOn ? <span>start: {p.startOn}</span> : null}
                      <span>due: {p.dueInDays ?? 30}d</span>
                      {p.graceDays ? <span>grace: {p.graceDays}d</span> : null}
                      <span>files: {Array.isArray(p.fileIds) ? p.fileIds.length : (p.fileId ? 1 : 0)}</span>
                      {(Array.isArray(p.fileIds) ? p.fileIds : (p.fileId ? [p.fileId] : [])).slice(0,1).map(fid => (
                        <a key={fid} href={`${apiBase}/api/files/${fid}`} target="_blank" rel="noreferrer">first file ↗</a>
                      ))}
                    </div>
                  </div>
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button className="btn ghost sm" onClick={() => del(p.id)}>Delete</button>
                  </div>
                </div>
              ))}
            </div>
          )
        )}
      </div>
    </div>
  );
};

export default Policies;
