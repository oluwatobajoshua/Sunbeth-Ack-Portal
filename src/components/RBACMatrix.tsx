import React, { useEffect, useMemo, useState } from 'react';
import { getPermissionCatalog, getRolePermissions, setRolePermissions, getUserPermissions, setUserPermissions, type PermissionDef } from '../services/rbacService';
import { getRoles, type DbRole } from '../services/dbService';
import { showToast } from '../utils/alerts';
import { isSQLiteEnabled } from '../utils/runtimeConfig';
import { useAuth } from '../context/AuthContext';
import { getUsers, getOrganizationStructure, type GraphUser } from '../services/graphUserService';

const KNOWN_ROLES = ['Admin','Manager','Employee'] as const;

const RBACMatrix: React.FC = () => {
  const sqlite = isSQLiteEnabled();
  const { account } = useAuth();
  const [perms, setPerms] = useState<PermissionDef[]>([]);
  const [roleMap, setRoleMap] = useState<Record<string, Record<string, boolean>>>({});
  const [busy, setBusy] = useState(false);
  const [tab, setTab] = useState<'role'|'user'>('role');

  const [userEmail, setUserEmail] = useState('');
  const [userMap, setUserMap] = useState<Record<string, boolean>>({});
  const [knownUsers, setKnownUsers] = useState<string[]>([]);
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState<{ department?: string; jobTitle?: string; location?: string }>({});
  const [org, setOrg] = useState<{ departments: string[]; jobTitles: string[]; locations: string[] }>({ departments: [], jobTitles: [], locations: [] });
  const [results, setResults] = useState<GraphUser[]>([]);
  const [searching, setSearching] = useState(false);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      try {
        const defs = await getPermissionCatalog();
        setPerms(defs);
        const next: Record<string, Record<string, boolean>> = {};
        for (const role of KNOWN_ROLES) {
          const rows = await getRolePermissions(role);
          const map: Record<string, boolean> = {};
          for (const d of defs) map[d.key] = false; // default false
          rows.forEach(r => { map[r.permKey] = !!r.value; });
          next[role] = map;
        }
        setRoleMap(next);
        // Load known users from roles DB
        if (sqlite) {
          const rows = await getRoles();
          const emails = Array.from(new Set(rows.map(r => r.email.toLowerCase())));
          setKnownUsers(emails);
        }
        // Load org structure for filters
        try {
          const token = await (async () => {
            try { return await (useAuth() as any).getToken?.(['User.Read.All']); } catch { return undefined; }
          })();
          if (token) {
            const o = await getOrganizationStructure(token);
            setOrg(o);
          }
        } catch {}
      } catch {
        showToast('Failed to load RBAC matrix', 'error');
      }
    })();
  }, [sqlite]);

  const saveRole = async (role: string) => {
    setBusy(true);
    try {
      await setRolePermissions(role, roleMap[role] || {});
      showToast(`Saved ${role} permissions`, 'success');
    } catch { showToast('Failed to save role permissions', 'error'); }
    finally { setBusy(false); }
  };

  const loadUser = async (email: string) => {
    setBusy(true);
    try {
      const rows = await getUserPermissions(email);
      const map: Record<string, boolean> = {};
      perms.forEach(p => map[p.key] = false);
      rows.forEach(r => map[r.permKey] = !!r.value);
      setUserMap(map);
    } catch { showToast('Failed to load user overrides', 'error'); }
    finally { setBusy(false); }
  };

  const saveUser = async () => {
    const e = userEmail.trim().toLowerCase();
    if (!e || !e.includes('@')) { showToast('Enter a valid user email', 'warning'); return; }
    setBusy(true);
    try { await setUserPermissions(e, userMap); showToast('Saved user overrides', 'success'); }
    catch { showToast('Failed to save user overrides', 'error'); }
    finally { setBusy(false); }
  };

  const runSearch = async () => {
    setErr(null);
    const q = search.trim();
    if (!q) { setResults([]); return; }
    setSearching(true);
    try {
      const token = await (useAuth() as any).getToken?.(['User.Read.All']);
      if (!token) throw new Error('Sign in to search directory');
      const list = await getUsers(token, { search: q, department: filters.department, jobTitle: filters.jobTitle, location: filters.location });
      setResults(Array.isArray(list) ? list.slice(0, 100) : []);
    } catch (e: any) {
      setErr(typeof e?.message === 'string' ? e.message : 'Search failed');
      setResults([]);
    } finally {
      setSearching(false);
    }
  };

  useEffect(() => {
    const t = setTimeout(() => { void runSearch(); }, 500);
    return () => clearTimeout(t);
  }, [search, filters.department, filters.jobTitle, filters.location]);

  const grouped = useMemo(() => {
    const m: Record<string, PermissionDef[]> = {};
    for (const p of perms) { const k = p.category || 'General'; (m[k] = m[k] || []).push(p); }
    return m;
  }, [perms]);

  return (
    <div>
      <div style={{ display: 'flex', gap: 8, marginBottom: 12, borderBottom: '1px solid #eee' }}>
        <button className={tab==='role'?'btn sm':'btn ghost sm'} onClick={() => setTab('role')}>By Role</button>
        <button className={tab==='user'?'btn sm':'btn ghost sm'} onClick={() => setTab('user')}>By User</button>
      </div>

      {tab === 'role' && (
        <div style={{ display: 'grid', gap: 16 }}>
          {Object.keys(grouped).map(cat => (
            <div key={cat}>
              <div style={{ fontWeight: 700, marginBottom: 6 }}>{cat}</div>
              <div style={{ overflowX: 'auto' }}>
                <table className="table" style={{ minWidth: 600 }}>
                  <thead>
                    <tr>
                      <th style={{ textAlign: 'left' }}>Permission</th>
                      {KNOWN_ROLES.map(r => (<th key={r} style={{ textAlign: 'center' }}>{r}</th>))}
                    </tr>
                  </thead>
                  <tbody>
                    {grouped[cat].map(p => (
                      <tr key={p.key}>
                        <td>
                          <div style={{ fontWeight: 500 }}>{p.label}</div>
                          <div className="small muted">{p.description}</div>
                        </td>
                        {KNOWN_ROLES.map(r => (
                          <td key={r} style={{ textAlign: 'center' }}>
                            <input
                              type="checkbox"
                              checked={!!roleMap[r]?.[p.key]}
                              onChange={e => setRoleMap(prev => ({ ...prev, [r]: { ...(prev[r] || {}), [p.key]: e.target.checked } }))}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ))}
          <div style={{ display: 'flex', gap: 8 }}>
            {KNOWN_ROLES.map(r => (
              <button key={r} className="btn sm" onClick={() => saveRole(r)} disabled={busy}>Save {r}</button>
            ))}
          </div>
        </div>
      )}

      {tab === 'user' && (
        <div style={{ display: 'grid', gap: 12 }}>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <input type="email" placeholder="user@domain.com" value={userEmail} onChange={e => setUserEmail(e.target.value)} />
            <button className="btn sm" onClick={() => loadUser(userEmail)} disabled={busy}>Load</button>
            {knownUsers.length > 0 && (
              <select value={userEmail} onChange={e => { setUserEmail(e.target.value); void loadUser(e.target.value); }}>
                <option value="">Select known user…</option>
                {knownUsers.map(e => (<option key={e} value={e}>{e}</option>))}
              </select>
            )}
            {account?.username && (
              <button className="btn ghost sm" onClick={() => { setUserEmail(account.username!); void loadUser(account.username!); }}>Use my email</button>
            )}
          </div>
          {/* Directory search to find users quickly */}
          <div className="card" style={{ padding: 12 }}>
            <div className="small" style={{ marginBottom: 8 }}>Search directory</div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8 }}>
              <input placeholder="Search by name or email" value={search} onChange={e => setSearch(e.target.value)} />
              <button className="btn sm" onClick={runSearch} disabled={searching}>Search</button>
            </div>
            <div className="small" style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
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
            {err && <div className="small" style={{ color: '#d33', marginTop: 6 }}>{err}</div>}
            {searching && <div className="small muted" style={{ marginTop: 6 }}>Searching...</div>}
            {!searching && results.length > 0 && (
              <div style={{ maxHeight: 220, overflowY: 'auto', marginTop: 8, display: 'grid', gap: 6 }}>
                {results.map(u => {
                  const email = (u.mail || u.userPrincipalName || '').trim();
                  return (
                    <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8, alignItems: 'center' }}>
                      <div>
                        <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName || email || u.id}</div>
                        <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                      </div>
                      <button className="btn ghost sm" onClick={() => { setUserEmail(email); void loadUser(email); setTab('user'); }}>Load</button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
          {perms.length > 0 && (
            <div style={{ display: 'grid', gap: 12 }}>
              {Object.keys(grouped).map(cat => (
                <div key={cat}>
                  <div style={{ fontWeight: 700, marginBottom: 6 }}>{cat}</div>
                  <div style={{ display: 'grid', gap: 8 }}>
                    {grouped[cat].map(p => (
                      <label key={p.key} className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <input type="checkbox" checked={!!userMap[p.key]} onChange={e => setUserMap(prev => ({ ...prev, [p.key]: e.target.checked }))} />
                        <span>{p.label}</span>
                        <span className="muted">— {p.description}</span>
                      </label>
                    ))}
                  </div>
                </div>
              ))}
              <div>
                <button className="btn sm" onClick={saveUser} disabled={busy || !userEmail}>Save User Overrides</button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default RBACMatrix;
