import React, { useEffect, useState } from 'react';
import { useRBAC } from '../context/RBACContext';
import { getApiBase } from '../utils/runtimeConfig';
import { Link } from 'react-router-dom';

type ModuleInfo = {
  name: string;
  title: string;
  version: string;
  adminRoute: string;
  routeBase: string;
  featureFlag: string;
  homeRoute?: string;
  enabled: boolean;
};

export default function ModulesHub() {
  const [mods, setMods] = useState<ModuleInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const rbac = useRBAC();
  const apiBase = (getApiBase() as string) || '';

  useEffect(() => {
    (async () => {
      setLoading(true);
      setError(null);
      try {
        // Prefer tenant-aware modules when available
        let j: any = null;
        try {
          const r1 = await fetch(`${apiBase}/api/tenant/modules`, { cache: 'no-store' });
          if (r1.ok) j = await r1.json();
        } catch {}
        if (!j || !Array.isArray(j?.modules)) {
          const r2 = await fetch(`${apiBase}/api/modules`, { cache: 'no-store' });
          j = await r2.json();
        }
        setMods(Array.isArray(j?.modules) ? j.modules : []);
      } catch (e) {
        setError('Failed to load modules');
      } finally {
        setLoading(false);
      }
    })();
  }, [apiBase]);

  const canSeeAdmin = rbac.canSeeAdmin || !!rbac.perms?.['viewAdmin'];
  const visible = mods.filter(m => m.enabled !== false);

  return (
    <div>
      <h2>Modules</h2>
      {loading && <div className="small muted">Loading modules…</div>}
      {error && <div className="small" style={{ color: 'crimson' }}>{error}</div>}
      {!loading && visible.length === 0 && (
        <div className="small muted">No modules are enabled for your account.</div>
      )}
      <div className="grid" style={{ gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))', gap: 16 }}>
        {visible.map(m => {
          const home = m.homeRoute || '/';
          return (
            <div key={m.name} className="card" style={{ padding: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
                <div style={{ fontWeight: 700 }}>{m.title}</div>
                <div className="small muted">v{m.version}</div>
              </div>
              <div className="small muted" style={{ margin: '6px 0 12px 0' }}>{m.name}</div>
              <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                <Link to={home} className="btn sm">Open</Link>
                {canSeeAdmin && (
                  <Link to={m.adminRoute} className="btn ghost sm">Admin</Link>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
