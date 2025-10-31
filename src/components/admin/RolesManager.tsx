/* eslint-disable max-lines-per-function, complexity */
import React, { useCallback, useEffect, useState } from 'react';
import { useAuth as useAuthCtx } from '../../context/AuthContext';
import { getRoles, createRole, deleteRole, type DbRole } from '../../services/dbService';
import { GraphUser, getUsers, getOrganizationStructure } from '../../services/graphUserService';
import { showToast } from '../../utils/alerts';
import { isSQLiteEnabled } from '../../utils/runtimeConfig';
import { getRolePermissions } from '../../services/rbacService';

const RolesManager: React.FC<{ canEdit: boolean; isSuperAdmin: boolean }> = ({ canEdit /*, isSuperAdmin*/ }) => {
  const { getToken, login, account } = useAuthCtx();
  const [roles, setRoles] = useState<DbRole[]>([]);
  const [busy, setBusy] = useState(false);
  const [email, setEmail] = useState('');
  const [role, setRole] = useState<string>('Manager');
  const [useCustomRole, setUseCustomRole] = useState(false);
  const [customRole, setCustomRole] = useState('');
  const sqliteEnabled = isSQLiteEnabled();

  // User search via Microsoft Graph
  const [userQuery, setUserQuery] = useState('');
  const [userResults, setUserResults] = useState<GraphUser[]>([]);
  const [userLoading, setUserLoading] = useState(false);
  const [userError, setUserError] = useState<string | null>(null);
  const [filters, setFilters] = useState<{ department?: string; jobTitle?: string; location?: string }>({});
  const [org, setOrg] = useState<{ departments: string[]; jobTitles: string[]; locations: string[] }>({ departments: [], jobTitles: [], locations: [] });
  const [selected, setSelected] = useState<Set<string>>(new Set());

  // Available roles from backend RBAC + defaults
  const [availableRoles, setAvailableRoles] = useState<string[]>(['SuperAdmin','Admin','Manager']);
  useEffect(() => {
    (async () => {
      try {
        const rp = await getRolePermissions().catch(() => []);
        const dyn = Array.from(new Set((rp as any[]).map((r: any) => String(r.role)).filter(Boolean)));
        const base = ['SuperAdmin','Admin','Manager'];
        const merged = Array.from(new Set([...base, ...dyn])).sort();
        setAvailableRoles(merged);
      } catch { /* ignore */ }
    })();
  }, []);

  const load = useCallback(async () => {
    if (!sqliteEnabled) { setRoles([]); return; }
    try {
      const list = await getRoles();
      setRoles(Array.isArray(list) ? list : []);
    } catch {
      setRoles([]);
    }
  }, [sqliteEnabled]);
  useEffect(() => { load(); }, [sqliteEnabled, load]);

  // Load organization structure for filters
  useEffect(() => {
    (async () => {
      try {
        const token = await getToken(['User.Read.All']);
        if (!token) return;
        const o = await getOrganizationStructure(token);
        setOrg(o);
      } catch { /* ignore */ }
    })();
  }, [getToken]);

  const searchUsers = useCallback(async () => {
    setUserError(null);
    if (!userQuery.trim()) { setUserResults([]); return; }
    setUserLoading(true);
    try {
      const token = await getToken(['User.Read.All']);
      if (!token) throw new Error('Sign-in required');
      const results = await getUsers(token, { search: userQuery.trim(), department: filters.department, jobTitle: filters.jobTitle, location: filters.location });
      setUserResults(Array.isArray(results) ? results.slice(0, 200) : []);
    } catch (e: any) {
      setUserError(typeof e?.message === 'string' ? e.message : 'Failed to search users');
      setUserResults([]);
    } finally {
      setUserLoading(false);
    }
  }, [filters.department, filters.jobTitle, filters.location, getToken, userQuery]);

  // Debounce search on inputs
  useEffect(() => {
    const t = setTimeout(() => { void searchUsers(); }, 450);
    return () => clearTimeout(t);
  }, [userQuery, filters.department, filters.jobTitle, filters.location, searchUsers]);

  const add = async () => {
    if (!canEdit || !sqliteEnabled) return;
    const e = email.trim().toLowerCase();
    if (!e || !e.includes('@')) { showToast('Enter a valid email', 'warning'); return; }
    const finalRole = useCustomRole ? (customRole.trim() || role) : role;
    if (!finalRole) { showToast('Choose a role or enter custom role', 'warning'); return; }
    setBusy(true);
    try {
      await createRole(e, finalRole);
      setEmail('');
      setCustomRole(''); setUseCustomRole(false);
      await load();
      showToast(`Role added (${finalRole})`, 'success');
    } catch {
      showToast('Failed to add role', 'error');
    } finally { setBusy(false); }
  };

  const assignToUser = async (u: GraphUser, r: string) => {
    if (!canEdit || !sqliteEnabled) return;
    const addr = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
    if (!addr) { showToast('User has no email/UPN', 'warning'); return; }
    setBusy(true);
    try {
      // If user already has a role, replace it when different
      const existing = roles.find(x => (x.email || '').toLowerCase() === addr);
      if (existing) {
        if (existing.role === r) {
          showToast(`${u.displayName || addr} already ${r}`, 'info');
          setBusy(false); return;
        }
        try { await deleteRole(existing.id); } catch { /* ignore */ }
      }
      await createRole(addr, r);
      await load();
      showToast(`Assigned ${r} to ${u.displayName || addr}`, 'success');
    } catch {
      showToast('Failed to assign role', 'error');
    } finally { setBusy(false); }
  };

  const assignBulk = async (r: string) => {
    if (!canEdit || !sqliteEnabled || selected.size === 0) return;
    setBusy(true);
    try {
      for (const id of Array.from(selected)) {
        const u = userResults.find(x => x.id === id);
        if (u) { await assignToUser(u, r); }
      }
      setSelected(new Set());
      showToast(`Assigned ${r} to selected user(s)`, 'success');
    } catch {
      showToast('Bulk assign failed', 'error');
    } finally { setBusy(false); }
  };

  const toggleSel = (id: string) => {
    setSelected(prev => { const n = new Set(prev); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  };

  const exportRolesCsv = () => {
    const rows = [['email','role']].concat(roles.map(r => [r.email, r.role]));
    const csv = rows.map(r => r.map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = 'roles.csv'; a.click(); URL.revokeObjectURL(url);
  };

  const remove = async (id: number) => {
    if (!canEdit || !sqliteEnabled) return;
    setBusy(true);
    try { await deleteRole(id); await load(); showToast('Role removed', 'success'); }
    catch { showToast('Failed to remove role', 'error'); }
    finally { setBusy(false); }
  };

  const grouped = roles.reduce((acc: Record<string, DbRole[]>, r) => {
    const key = r.role || 'Unknown';
    (acc[key] = acc[key] || []).push(r);
    return acc;
  }, {} as Record<string, DbRole[]>);

  if (!sqliteEnabled) return <div className="small muted">Enable SQLite to manage roles.</div>;
  return (
    <div>
      {/* Directory user search and quick-assign */}
      <div className="card" style={{ padding: 12, marginBottom: 12 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
          <div>
            <div style={{ fontWeight: 700 }}>Find users</div>
            <div className="small muted">Search your directory and assign Admin/Manager</div>
          </div>
          {!account && (
            <button className="btn ghost sm" onClick={() => login()}>Sign in</button>
          )}
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8, marginTop: 8 }}>
          <input
            placeholder="Search by name or email"
            value={userQuery}
            onChange={e => setUserQuery(e.target.value)}
          />
          <button className="btn sm" onClick={searchUsers} disabled={userLoading}>Search</button>
        </div>
        {/* Filters */}
        <div className="small" style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
          <select value={filters.department || ''} onChange={e => setFilters(f => ({ ...f, department: e.target.value || undefined }))}>
            <option value="">All departments</option>
            {org.departments.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={filters.jobTitle || ''} onChange={e => setFilters(f => ({ ...f, jobTitle: e.target.value || undefined }))}>
            <option value="">All job titles</option>
            {org.jobTitles.map(j => <option key={j} value={j}>{j}</option>)}
          </select>
          <select value={filters.location || ''} onChange={e => setFilters(f => ({ ...f, location: e.target.value || undefined }))}>
            <option value="">All locations</option>
            {org.locations.map(l => <option key={l} value={l}>{l}</option>)}
          </select>
        </div>
        {/* Bulk actions */}
        {selected.size > 0 && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 8, flexWrap: 'wrap' }}>
            <span>{selected.size} selected</span>
            <select onChange={e => setRole(e.target.value)} value={role} style={{ padding: 6, border: '1px solid #ddd', borderRadius: 6 }}>
              {availableRoles.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
            <button className="btn ghost sm" onClick={() => assignBulk(role)} disabled={!canEdit || busy}>Assign Role</button>
            <button className="btn ghost sm" onClick={() => assignBulk('SuperAdmin')} disabled={!canEdit || busy}>Assign SuperAdmin</button>
          </div>
        )}
        {userError && <div className="small" style={{ color: '#d33', marginTop: 6 }}>{userError}</div>}
        {userLoading && <div className="small muted" style={{ marginTop: 6 }}>Loading...</div>}
        {!userLoading && userResults.length > 0 && (
          <div style={{ marginTop: 8, maxHeight: 220, overflowY: 'auto', display: 'grid', gap: 6 }}>
            {userResults.map(u => {
              const email = (u.mail || u.userPrincipalName || '').trim();
              const existing = roles.find(r => (r.email || '').toLowerCase() === (email || '').toLowerCase());
              return (
                <div key={u.id} style={{ display: 'grid', gridTemplateColumns: 'auto 1fr auto auto', gap: 8, alignItems: 'center' }}>
                  <input type="checkbox" checked={selected.has(u.id)} onChange={() => toggleSel(u.id)} />
                  <div>
                    <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName || email || u.id}</div>
                    <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                    {existing && <span className="badge" style={{ marginTop: 4 }}>{existing.role}</span>}
                  </div>
                  {!existing && (
                    <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                      <select value={role} onChange={e => setRole(e.target.value)} style={{ padding: 6, border: '1px solid #ddd', borderRadius: 6 }}>
                        {availableRoles.map(r => <option key={r} value={r}>{r}</option>)}
                      </select>
                      <button className="btn ghost sm" onClick={() => assignToUser(u, role)} disabled={!canEdit || busy}>Assign</button>
                    </div>
                  )}
                  {existing && (
                    <div style={{ display: 'flex', gap: 6 }}>
                      <select value={role} onChange={e => setRole(e.target.value)} style={{ padding: 6, border: '1px solid #ddd', borderRadius: 6 }}>
                        {availableRoles.map(r => <option key={r} value={r}>{r}</option>)}
                      </select>
                      <button className="btn ghost sm" onClick={() => assignToUser(u, role)} disabled={!canEdit || busy}>Change to {role}</button>
                      <button className="btn ghost sm" onClick={async () => { try { setBusy(true); await deleteRole(existing.id); await load(); showToast('Role removed', 'success'); } catch { showToast('Failed to remove role', 'error'); } finally { setBusy(false); } }} disabled={!canEdit || busy}>Remove</button>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap', marginBottom: 12 }}>
        <input type="email" placeholder="user@domain.com" value={email} onChange={e => setEmail(e.target.value)} style={{ padding: 8, border: '1px solid #ddd', borderRadius: 6 }} />
        <select value={useCustomRole ? '::custom' : role} onChange={e => {
          const v = e.target.value; if (v === '::custom') setUseCustomRole(true); else { setUseCustomRole(false); setRole(v); }
        }} style={{ padding: 8, border: '1px solid #ddd', borderRadius: 6 }}>
          {availableRoles.map(r => <option key={r} value={r}>{r}</option>)}
          <option value="::custom">Custom…</option>
        </select>
        {useCustomRole && (
          <input placeholder="Custom role" value={customRole} onChange={e => setCustomRole(e.target.value)} style={{ padding: 8, border: '1px solid #ddd', borderRadius: 6, minWidth: 160 }} />
        )}
        <button className="btn sm" onClick={add} disabled={!canEdit || busy}>Add</button>
        {!canEdit && <span className="small muted">Read-only</span>}
        <button className="btn ghost sm" onClick={exportRolesCsv} title="Export current role assignments as CSV">Export CSV</button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: 12 }}>
        {Object.keys(grouped).sort().map(k => (
          <div key={k} className="card" style={{ padding: 12 }}>
            <div style={{ fontWeight: 700, marginBottom: 6 }}>{k}s</div>
            {Array.isArray(grouped[k]) && grouped[k].length > 0 ? (
              <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                {grouped[k].map(r => (
                  <li key={r.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 0', borderBottom: '1px solid #f2f2f2' }}>
                    <span className="small">{r.email}</span>
                    {canEdit && <button className="btn ghost sm" onClick={() => remove(r.id)} disabled={busy}>Remove</button>}
                  </li>
                ))}
              </ul>
            ) : (
              <div className="small muted">No {k.toLowerCase()}s assigned</div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};

export default RolesManager;
