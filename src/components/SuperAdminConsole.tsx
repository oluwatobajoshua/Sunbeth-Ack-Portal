import React, { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { getApiBase } from '../utils/runtimeConfig';
import { useRBAC } from '../context/RBACContext';

type Tenant = { id: number; name: string; code: string; isActive: boolean; isOwner: boolean; parentId: number|null; modulesEnabled: number; activeLicenses: number };
type ModuleInfo = { name: string; title: string; version: string; adminRoute: string; routeBase: string; featureFlag: string; enabled: boolean };
type Domain = { id: number; domain: string; isPrimary: boolean; verified: boolean; addedAt?: string };
type Theme = { darkMode?: boolean; logoUrl?: string; colors?: Partial<{ primary: string; accent: string; bg: string; bgElevated: string; card: string; muted: string }>; cssVars?: Record<string,string> };
type CustomizationRequest = { id: number; tenantId: number; tenantName: string; contactName?: string; contactEmail?: string; contactPhone?: string; description: string; scope?: string; priority: string; status: string; createdAt: string };
type ThemeSummary = { id: number; name: string; description?: string | null; baseThemeId?: number | null; isSystem?: boolean; createdAt?: string; updatedAt?: string };
export default function SuperAdminConsole() {
  const [adminEmail, setAdminEmail] = useState<string>(() => {
    try { return localStorage.getItem('sunbeth_superadmin_email') || ''; } catch { return ''; }
  });
  const [tenants, setTenants] = useState<Tenant[]>([]);
  const [selected, setSelected] = useState<number | null>(null);
  const [mods, setMods] = useState<ModuleInfo[]>([]);
  const [domains, setDomains] = useState<Domain[]>([]);
  const [theme, setTheme] = useState<Theme>({});
  const [custReqs, setCustReqs] = useState<CustomizationRequest[]>([]);
  const [newDomain, setNewDomain] = useState<string>('');
  const [newDomainPrimary, setNewDomainPrimary] = useState<boolean>(false);
  const [themes, setThemes] = useState<ThemeSummary[]>([]);
  const [selectedThemeId, setSelectedThemeId] = useState<number | null>(null);
  const [themeEditorLight, setThemeEditorLight] = useState<Theme>({});
  const [themeEditorDark, setThemeEditorDark] = useState<Theme>({ darkMode: true });
  const [themeEditorName, setThemeEditorName] = useState<string>('');
  const [assignModule, setAssignModule] = useState<string>('');
  const [assignPlugin, setAssignPlugin] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  // Roles & Settings
  const [roleUsers, setRoleUsers] = useState<Array<{ email: string; roles: string[] }>>([]);
  const [newRoleEmail, setNewRoleEmail] = useState('');
  const [newRoleRole, setNewRoleRole] = useState('SuperAdmin');
  const [settings, setSettings] = useState<Record<string, any>>({});
  // Feature Flags
  const [globalFlags, setGlobalFlags] = useState<Record<string, boolean>>({});
  const [tenantFlags, setTenantFlags] = useState<Record<string, boolean>>({});
  const [newFlagKey, setNewFlagKey] = useState('');
  const rbac = useRBAC();
  const apiBase = (getApiBase() as string) || '';

  // Helper to append adminEmail to /api/admin calls without requiring env changes
  const adminUrl = (path: string) => {
    const url = `${apiBase}${path}`;
    if (!adminEmail) return url;
    return url + (url.includes('?') ? `&adminEmail=${encodeURIComponent(adminEmail)}` : `?adminEmail=${encodeURIComponent(adminEmail)}`);
  };

  const loadTenants = async () => {
    setLoading(true); setError(null);
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants`));
      const j = await res.json();
      setTenants(Array.isArray(j?.tenants) ? j.tenants : []);
    } catch (e) {
      setError('Failed to load tenants');
    } finally { setLoading(false); }
  };
  const loadMods = async (tenantId: number) => {
    setError(null); setMods([]);
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${tenantId}/modules`));
      const j = await res.json();
      setMods(Array.isArray(j?.modules) ? j.modules : []);
    } catch (e) { setError('Failed to load tenant modules'); }
  };
  const loadDomains = async (tenantId: number) => {
    setError(null); setDomains([]);
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${tenantId}/domains`));
      const j = await res.json();
      setDomains(Array.isArray(j?.domains) ? j.domains : []);
    } catch { setError('Failed to load domains'); }
  };
  const loadTheme = async (tenantId: number) => {
    setError(null); setTheme({});
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${tenantId}/theme`));
      const j = await res.json();
      setTheme((j && j.theme) ? j.theme : {});
    } catch { setTheme({}); }
  };
  const loadCustomizationRequests = async (tenantId: number) => {
    setError(null); setCustReqs([]);
    try {
      const res = await fetch(adminUrl(`/api/admin/customization-requests?tenantId=${tenantId}`));
      const j = await res.json();
      setCustReqs(Array.isArray(j?.requests) ? j.requests : []);
    } catch { /* ignore */ }
  };
  const loadThemeLibrary = async () => {
    try {
      const res = await fetch(adminUrl(`/api/admin/themes`));
      const j = await res.json();
      setThemes(Array.isArray(j?.themes) ? j.themes : []);
    } catch { /* ignore */ }
  };
  const loadRoles = async () => {
    try {
      const res = await fetch(adminUrl(`/api/admin/roles`));
      const j = await res.json();
      setRoleUsers(Array.isArray(j?.users) ? j.users : []);
    } catch {}
  };
  const addRole = async () => {
    const e = newRoleEmail.trim().toLowerCase();
    if (!e || !newRoleRole) return;
    try {
      await fetch(adminUrl(`/api/admin/roles`), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: e, role: newRoleRole }) });
      setNewRoleEmail('');
      await loadRoles();
    } catch {}
  };
  const removeRole = async (email: string, role: string) => {
    try {
      await fetch(adminUrl(`/api/admin/roles?email=${encodeURIComponent(email)}&role=${encodeURIComponent(role)}`), { method: 'DELETE' });
      await loadRoles();
    } catch {}
  };
  const loadSettings = async () => {
    try {
      const res = await fetch(adminUrl(`/api/admin/settings`));
      const j = await res.json();
      setSettings(j?.settings || {});
    } catch {}
  };
  const saveSettings = async () => {
    try {
      await fetch(adminUrl(`/api/admin/settings`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ settings }) });
    } catch {}
  };
  // Feature Flags helpers
  const RECOMMENDED_FLAGS = [
    { key: 'ff_workflows_enabled', label: 'Workflows/Approvals' },
    { key: 'ff_analytics_v2', label: 'Analytics v2' },
    { key: 'ff_external_users', label: 'External Users' },
    { key: 'ff_extensions_marketplace', label: 'Extensions Marketplace' },
    { key: 'ff_document_e_signature', label: 'Document e-signature' }
  ];
  const loadGlobalFlags = async () => {
    try {
      const res = await fetch(adminUrl(`/api/admin/feature-flags`));
      const j = await res.json();
      setGlobalFlags(j?.flags || {});
    } catch {}
  };
  const saveGlobalFlags = async () => {
    try {
      await fetch(adminUrl(`/api/admin/feature-flags`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ flags: globalFlags }) });
    } catch {}
  };
  const loadTenantFlags = async (tenantId: number) => {
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${tenantId}/feature-flags`));
      const j = await res.json();
      setTenantFlags(j?.flags || {});
    } catch {}
  };
  const saveTenantFlags = async () => {
    if (selected == null) return;
    try {
      await fetch(adminUrl(`/api/admin/tenants/${selected}/feature-flags`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ flags: tenantFlags }) });
    } catch {}
  };
  const openTheme = async (id: number) => {
    try {
      const res = await fetch(adminUrl(`/api/admin/themes/${id}`));
      const t = await res.json();
      setSelectedThemeId(id);
      setThemeEditorName(t.name || '');
      setThemeEditorLight(t.light || {});
      setThemeEditorDark(t.dark || { darkMode: true });
    } catch { /* ignore */ }
  };
  const newTheme = async () => {
    try {
      const res = await fetch(adminUrl(`/api/admin/themes`), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: 'New Theme', light: {}, dark: { darkMode: true } }) });
      const j = await res.json();
      if (j?.id) { await loadThemeLibrary(); await openTheme(j.id); }
    } catch {}
  };
  const cloneTheme = async (id: number) => {
    try {
      const res = await fetch(adminUrl(`/api/admin/themes/${id}/clone`), { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: `${themes.find(x=>x.id===id)?.name || 'Theme'} Copy` }) });
      const j = await res.json();
      if (j?.id) { await loadThemeLibrary(); await openTheme(j.id); }
    } catch {}
  };
  const deleteTheme = async (id: number) => {
    try {
      await fetch(adminUrl(`/api/admin/themes/${id}`), { method: 'DELETE' });
      if (selectedThemeId === id) setSelectedThemeId(null);
      await loadThemeLibrary();
    } catch {}
  };
  const saveThemeEditor = async () => {
    if (selectedThemeId == null) return;
    try {
      await fetch(adminUrl(`/api/admin/themes/${selectedThemeId}`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: themeEditorName, light: themeEditorLight, dark: themeEditorDark }) });
      await loadThemeLibrary();
    } catch {}
  };
  const assignTheme = async (targetType: 'tenant'|'module'|'plugin'|'global') => {
    if (selectedThemeId == null) return;
    try {
      const targetId = targetType==='tenant' ? String(selected) : (targetType==='module' ? assignModule : (targetType==='plugin' ? assignPlugin : null));
      const payload: any = { themeId: selectedThemeId, targetType, targetId };
      await fetch(adminUrl(`/api/admin/theme-assignments`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      // refresh effective theme for tenant panel
      if (selected != null) await loadTheme(selected);
    } catch {}
  };

  useEffect(() => { loadTenants(); }, []);
  useEffect(() => {
    if (selected != null) {
      loadMods(selected);
      loadDomains(selected);
      loadTheme(selected);
      loadCustomizationRequests(selected);
      loadTenantFlags(selected);
    }
  }, [selected]);
  useEffect(() => { loadThemeLibrary(); loadRoles(); loadSettings(); loadGlobalFlags(); }, []);

  const toggleModule = async (m: ModuleInfo, enabled: boolean) => {
    if (selected == null) return;
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${selected}/modules`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ module: m.name, enabled }) });
      if (!res.ok) throw new Error('save_failed');
      setMods(prev => prev.map(x => x.name === m.name ? { ...x, enabled } : x));
    } catch { setError('Failed to save'); }
  };
  const addDomain = async () => {
    if (selected == null || !newDomain.trim()) return;
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${selected}/domains`), {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ domain: newDomain.trim(), isPrimary: !!newDomainPrimary })
      });
      if (!res.ok) throw new Error('create_failed');
      setNewDomain(''); setNewDomainPrimary(false);
      await loadDomains(selected);
    } catch { setError('Failed to add domain'); }
  };
  const deleteDomain = async (domainId: number) => {
    if (selected == null) return;
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${selected}/domains/${domainId}`), { method: 'DELETE' });
      if (!res.ok) throw new Error('delete_failed');
      await loadDomains(selected);
    } catch { setError('Failed to delete domain'); }
  };
  const saveTheme = async () => {
    if (selected == null) return;
    try {
      const res = await fetch(adminUrl(`/api/admin/tenants/${selected}/theme`), { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ theme }) });
      if (!res.ok) throw new Error('save_failed');
    } catch { setError('Failed to save theme'); }
  };

  const primaryColor = theme?.colors?.primary || '';
  const accentColor = theme?.colors?.accent || '';
  const onChangePrimary = (v: string) => setTheme(prev => ({ ...prev, colors: { ...(prev.colors || {}), primary: v || undefined } }));
  const onChangeAccent = (v: string) => setTheme(prev => ({ ...prev, colors: { ...(prev.colors || {}), accent: v || undefined } }));
  const onChangeLogo = (v: string) => setTheme(prev => ({ ...prev, logoUrl: v || undefined }));
  const onToggleDark = (v: boolean) => setTheme(prev => ({ ...prev, darkMode: v }));
  const themePreviewStyle = useMemo<React.CSSProperties>(() => ({
    display: 'inline-flex', gap: 8, alignItems: 'center', padding: 8, border: '1px solid #eee', borderRadius: 8
  }), []);

  if (!rbac.canSeeAdmin) return <div className="small muted">Access denied</div>;

  return (
    <div>
      <h2>Super Admin Console</h2>
      <div className="card" style={{ padding: 12, marginBottom: 12 }}>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <label className="small" title="Used to authorize admin API calls without env changes">
            Super Admin Email
            <input
              type="email"
              placeholder="admin@example.com"
              value={adminEmail}
              onChange={e => { const v = e.target.value; setAdminEmail(v); try { localStorage.setItem('sunbeth_superadmin_email', v); } catch {} }}
              style={{ marginLeft: 8, padding: 8, borderRadius: 6, border: '1px solid #e6e6e6', minWidth: 260 }}
            />
          </label>
          <div className="small muted">This value is appended as <code>adminEmail</code> to admin API calls.</div>
        </div>
      </div>
      {/* Global Settings & Roles */}
      <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
        <div className="card" style={{ padding: 12 }}>
          <div style={{ fontWeight: 700, marginBottom: 10 }}>Global Settings</div>
          <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <label className="small" title="Enable external users / MFA / password reset endpoints">
              External Support Enabled
              <input type="checkbox" checked={!!settings.external_support_enabled} onChange={e => setSettings(s => ({ ...s, external_support_enabled: e.target.checked }))} style={{ marginLeft: 8 }} />
            </label>
            <label className="small" title="Comma-separated list of allowed origins for CORS (e.g., http://localhost:3000,https://portal.example.com)">
              Allowed Origins
              <input value={settings.allowed_origins || ''} onChange={e => setSettings(s => ({ ...s, allowed_origins: e.target.value }))} placeholder="http://localhost:3000" style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
            </label>
            <label className="small" title="If set, supports CODE.baseDomain routing when subdomains are enabled">
              Tenant Base Domain
              <input value={settings.tenant_base_domain || ''} onChange={e => setSettings(s => ({ ...s, tenant_base_domain: e.target.value }))} placeholder="local.test" style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
            </label>
            <label className="small">
              Subdomain Routing Enabled
              <input type="checkbox" checked={!!settings.tenant_subdomain_enabled} onChange={e => setSettings(s => ({ ...s, tenant_subdomain_enabled: e.target.checked }))} style={{ marginLeft: 8 }} />
            </label>
            <label className="small" title="Base URL used in onboarding links when origin is not available">
              Frontend Base URL
              <input value={settings.frontend_base_url || ''} onChange={e => setSettings(s => ({ ...s, frontend_base_url: e.target.value }))} placeholder="https://portal.example.com" style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
            </label>
          </div>
          <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 12 }}>
            <button className="btn sm" onClick={saveSettings}>Save Settings</button>
          </div>
        </div>
        <div className="card" style={{ padding: 12 }}>
          <div style={{ fontWeight: 700, marginBottom: 10 }}>Roles</div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
            <input type="email" placeholder="user@example.com" value={newRoleEmail} onChange={e => setNewRoleEmail(e.target.value)} style={{ flex: 1, padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
            <select value={newRoleRole} onChange={e => setNewRoleRole(e.target.value)} style={{ padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }}>
              <option value="SuperAdmin">SuperAdmin</option>
              <option value="Admin">Admin</option>
              <option value="Manager">Manager</option>
            </select>
            <button className="btn sm" onClick={addRole} disabled={!newRoleEmail.trim()}>Add</button>
          </div>
          <div className="grid" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 12 }}>
            {roleUsers.map(u => (
              <div key={u.email} className="card" style={{ padding: 10 }}>
                <div style={{ fontWeight: 700 }}>{u.email}</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 8 }}>
                  {u.roles.map(r => (
                    <span key={r} className="badge" style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                      {r}
                      <button className="btn ghost xs" onClick={() => removeRole(u.email, r)} title="Remove role">×</button>
                    </span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
      {/* Feature Flags */}
      <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
        <div className="card" style={{ padding: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ fontWeight: 700 }}>Global Feature Flags</div>
            <button className="btn sm" onClick={saveGlobalFlags}>Save</button>
          </div>
          <div style={{ display: 'grid', gap: 8, marginTop: 10 }}>
            {RECOMMENDED_FLAGS.map(f => (
              <label key={f.key} className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="checkbox" checked={!!globalFlags[f.key]} onChange={e => setGlobalFlags(g => ({ ...g, [f.key]: e.target.checked }))} /> {f.label}
                <span className="small muted">({f.key})</span>
              </label>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 12 }}>
            <input placeholder="ff_custom_flag" value={newFlagKey} onChange={e => setNewFlagKey(e.target.value)} style={{ flex: 1, padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
            <button className="btn ghost sm" onClick={() => { const k = newFlagKey.trim(); if (/^ff_[a-z0-9._-]+$/i.test(k)) { setGlobalFlags(g => ({ ...g, [k]: true })); setNewFlagKey(''); } }}>Add Flag</button>
          </div>
        </div>
        <div className="card" style={{ padding: 12 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ fontWeight: 700 }}>Tenant Feature Flags</div>
            <button className="btn sm" onClick={saveTenantFlags} disabled={selected==null}>Save</button>
          </div>
          {selected == null ? (
            <div className="small muted" style={{ marginTop: 10 }}>Select a tenant to view and override flags.</div>
          ) : (
            <div style={{ display: 'grid', gap: 8, marginTop: 10 }}>
              {RECOMMENDED_FLAGS.map(f => (
                <label key={f.key} className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <input type="checkbox" checked={!!tenantFlags[f.key]} onChange={e => setTenantFlags(g => ({ ...g, [f.key]: e.target.checked }))} /> {f.label}
                  <span className="small muted">({f.key})</span>
                </label>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="grid" style={{ gridTemplateColumns: '280px 1fr', gap: 16 }}>
        <div className="card" style={{ padding: 12 }}>
          <div style={{ fontWeight: 700, marginBottom: 8 }}>Tenants</div>
          {loading && <div className="small muted">Loading…</div>}
          {error && <div className="small" style={{ color: 'crimson' }}>{error}</div>}
          <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
            {tenants.map(t => (
              <li key={t.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 4px', borderBottom: '1px solid #eee', cursor: 'pointer', background: selected===t.id? '#f5f5f5': undefined }} onClick={() => setSelected(t.id)}>
                <div>
                  <div style={{ fontWeight: 600 }}>{t.name} {t.isOwner && <span className="small muted" title="Owner">(Owner)</span>}</div>
                  <div className="small muted">{t.code}</div>
                </div>
                <div className="small muted" title="Enabled modules">{t.modulesEnabled} mods</div>
              </li>
            ))}
          </ul>
        </div>
        <div>
          {selected == null ? (
            <div className="small muted">Select a tenant to manage entitlements.</div>
          ) : (
            <div style={{ display: 'grid', gap: 16 }}>
              <div className="card" style={{ padding: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ fontWeight: 700 }}>Modules</div>
                  <Link to={`/admin`} className="btn ghost sm">Global Admin</Link>
                </div>
                <div className="grid" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16, marginTop: 12 }}>
                  {mods.map(m => (
                    <div key={m.name} className="card" style={{ padding: 12 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
                        <div style={{ fontWeight: 700 }}>{m.title}</div>
                        <div className="small muted">v{m.version}</div>
                      </div>
                      <div className="small muted" style={{ margin: '4px 0 10px 0' }}>{m.name}</div>
                      <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                        <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                          <input type="checkbox" checked={!!m.enabled} onChange={e => toggleModule(m, e.target.checked)} />
                          <span>{m.enabled ? 'Enabled' : 'Disabled'}</span>
                        </label>
                        <Link to={m.adminRoute} className="btn ghost sm">Admin</Link>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="card" style={{ padding: 12 }}>
                <div style={{ fontWeight: 700, marginBottom: 10 }}>Custom Domains</div>
                <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                  <input placeholder="tenant.example.com" value={newDomain} onChange={e => setNewDomain(e.target.value)} style={{ flex: 1, padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                  <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <input type="checkbox" checked={newDomainPrimary} onChange={e => setNewDomainPrimary(e.target.checked)} /> Primary
                  </label>
                  <button className="btn sm" onClick={addDomain} disabled={!newDomain.trim()}>Add</button>
                </div>
                <div className="grid" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 12 }}>
                  {domains.map(d => (
                    <div key={d.id} className="card" style={{ padding: 10 }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <div>
                          <div style={{ fontWeight: 700 }}>{d.domain} {d.isPrimary && <span className="badge done" style={{ marginLeft: 6 }}>Primary</span>}</div>
                          <div className="small muted">{d.verified ? 'Verified' : 'Unverified'}</div>
                        </div>
                        <button className="btn ghost sm" onClick={() => deleteDomain(d.id)}>Delete</button>
                      </div>
                    </div>
                  ))}
                </div>
                <div className="small muted" style={{ marginTop: 8 }}>Note: Verification and SSL automation will be added in a later step.</div>
              </div>

              <div className="card" style={{ padding: 12 }}>
                <div style={{ fontWeight: 700, marginBottom: 10 }}>Theme</div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 12 }}>
                  <label className="small">Primary Color
                    <input type="text" placeholder="#0b5fff" value={primaryColor} onChange={e => onChangePrimary(e.target.value)} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                  </label>
                  <label className="small">Accent Color
                    <input type="text" placeholder="#ff5a0b" value={accentColor} onChange={e => onChangeAccent(e.target.value)} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                  </label>
                  <label className="small">Logo URL
                    <input type="text" placeholder="/images/logo.svg" value={theme?.logoUrl || ''} onChange={e => onChangeLogo(e.target.value)} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                  </label>
                </div>
                <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginTop: 10 }}>
                  <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <input type="checkbox" checked={!!theme.darkMode} onChange={e => onToggleDark(e.target.checked)} /> Dark mode
                  </label>
                  <div style={themePreviewStyle}>
                    <span className="small muted">Preview:</span>
                    <span className="badge" style={{ background: primaryColor || 'var(--primary)', color: '#fff' }}>Primary</span>
                    <span className="badge" style={{ background: accentColor || 'var(--accent)', color: '#fff' }}>Accent</span>
                  </div>
                  <div className="spacer" />
                  <button className="btn sm" onClick={saveTheme}>Save Theme</button>
                </div>
              </div>

              <div className="card" style={{ padding: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div style={{ fontWeight: 700 }}>Theme Library & Assignments</div>
                  <div style={{ display: 'flex', gap: 8 }}>
                    <button className="btn ghost sm" onClick={newTheme}>New Theme</button>
                    {selectedThemeId != null && (
                      <>
                        <button className="btn ghost sm" onClick={() => cloneTheme(selectedThemeId!)}>Clone</button>
                        <button className="btn ghost sm" onClick={() => deleteTheme(selectedThemeId!)}>Delete</button>
                      </>
                    )}
                  </div>
                </div>
                <div className="grid" style={{ gridTemplateColumns: '220px 1fr', gap: 12, marginTop: 10 }}>
                  <div>
                    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                      {themes.map(t => (
                        <li key={t.id} onClick={() => openTheme(t.id)} style={{ padding: '8px 6px', borderBottom: '1px solid #eee', cursor: 'pointer', background: selectedThemeId===t.id? '#f5f5f5': undefined }}>
                          <div style={{ fontWeight: 600 }}>{t.name}</div>
                          <div className="small muted">{t.description || '—'}</div>
                        </li>
                      ))}
                    </ul>
                  </div>
                  <div>
                    {selectedThemeId == null ? (
                      <div className="small muted">Select a theme to edit and assign.</div>
                    ) : (
                      <div style={{ display: 'grid', gap: 10 }}>
                        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                          <input value={themeEditorName} onChange={e => setThemeEditorName(e.target.value)} placeholder="Theme name" style={{ flex: 1, padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                          <button className="btn sm" onClick={saveThemeEditor}>Save Theme</button>
                        </div>
                        <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                          <div className="card" style={{ padding: 10 }}>
                            <div style={{ fontWeight: 700, marginBottom: 8 }}>Light Variant</div>
                            <label className="small">Primary
                              <input value={themeEditorLight?.colors?.primary || ''} onChange={e => setThemeEditorLight(prev => ({ ...prev, colors: { ...(prev.colors||{}), primary: e.target.value || undefined } }))} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                            </label>
                            <label className="small">Accent
                              <input value={themeEditorLight?.colors?.accent || ''} onChange={e => setThemeEditorLight(prev => ({ ...prev, colors: { ...(prev.colors||{}), accent: e.target.value || undefined } }))} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                            </label>
                          </div>
                          <div className="card" style={{ padding: 10 }}>
                            <div style={{ fontWeight: 700, marginBottom: 8 }}>Dark Variant</div>
                            <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                              <input type="checkbox" checked={!!themeEditorDark.darkMode} onChange={e => setThemeEditorDark(prev => ({ ...prev, darkMode: e.target.checked }))} /> Enable Dark Mode
                            </label>
                            <label className="small">Primary
                              <input value={themeEditorDark?.colors?.primary || ''} onChange={e => setThemeEditorDark(prev => ({ ...prev, colors: { ...(prev.colors||{}), primary: e.target.value || undefined } }))} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                            </label>
                            <label className="small">Accent
                              <input value={themeEditorDark?.colors?.accent || ''} onChange={e => setThemeEditorDark(prev => ({ ...prev, colors: { ...(prev.colors||{}), accent: e.target.value || undefined } }))} style={{ width: '100%', padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                            </label>
                          </div>
                        </div>
                        <div className="card" style={{ padding: 10 }}>
                          <div style={{ fontWeight: 700, marginBottom: 8 }}>Assign Theme</div>
                          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                            <button className="btn sm" onClick={() => assignTheme('tenant')} disabled={selected==null}>Assign to Tenant</button>
                            <select value={assignModule} onChange={e => setAssignModule(e.target.value)} style={{ padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }}>
                              <option value="">Select Module…</option>
                              {mods.map(m => (<option key={m.name} value={m.name}>{m.title}</option>))}
                            </select>
                            <button className="btn ghost sm" onClick={() => assignTheme('module')} disabled={!assignModule}>Assign to Module</button>
                            <input value={assignPlugin} onChange={e => setAssignPlugin(e.target.value)} placeholder="plugin-id" style={{ padding: 8, borderRadius: 6, border: '1px solid #e6e6e6' }} />
                            <button className="btn ghost sm" onClick={() => assignTheme('plugin')} disabled={!assignPlugin}>Assign to Plugin</button>
                            <button className="btn ghost sm" onClick={() => assignTheme('global')}>Assign as Global</button>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="card" style={{ padding: 12 }}>
                <div style={{ fontWeight: 700, marginBottom: 10 }}>Customization Requests</div>
                {custReqs.length === 0 ? (
                  <div className="small muted">No requests yet.</div>
                ) : (
                  <div className="grid" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: 12 }}>
                    {custReqs.map(r => (
                      <div key={r.id} className="card" style={{ padding: 10 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <div style={{ fontWeight: 700 }}>{r.scope || 'General'} <span className="small muted">· {r.priority}</span></div>
                          <div className="badge not">{r.status}</div>
                        </div>
                        <div className="small" style={{ marginTop: 6 }}>{r.description}</div>
                        <div className="small muted" style={{ marginTop: 8 }}>{r.contactName || r.contactEmail || '—'} · {new Date(r.createdAt).toLocaleString()}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
    
  );
}
