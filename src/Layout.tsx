import React, { useEffect, useRef, useState } from 'react';
import { useLocation, useNavigate, Link } from 'react-router-dom';
import { useAuth } from './context/AuthContext';
import { useExternalAuth } from './context/ExternalAuthContext';
import { useFeatureFlags } from './context/FeatureFlagsContext';
import { useRBAC } from './context/RBACContext';
import { useTenant } from './context/TenantContext';
// import DevPanel from './components/DevPanel';
import { info } from './diagnostics/logger';
import { getBatches, getUserProgress } from './services/dbService';
import DancingLogoOverlay from './components/DancingLogoOverlay';
import { enforceDuePolicies } from './utils/policiesDue';

const Layout: React.FC<React.PropsWithChildren> = ({ children }) => {
  const { account, token, photo, login, logout } = useAuth();
  const { user: externalUser, isAuthenticated: isExternal, logout: externalLogout } = useExternalAuth();
  const { externalSupport, loaded: flagsLoaded } = useFeatureFlags();
  const { tenant } = useTenant();
  const [theme, setTheme] = useState<'light'|'dark'>(() => {
    try {
      const ls = localStorage.getItem('sunbeth_theme');
      if (ls === 'light' || ls === 'dark') return ls;
    } catch { /* ignore */ }
    try {
      const attr = document.documentElement.getAttribute('data-theme');
      if (attr === 'light' || attr === 'dark') return attr;
    } catch { /* ignore */ }
    try { return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'; } catch { return 'light'; }
  });
  const [stickyHeader, setStickyHeader] = useState<boolean>(() => {
    try { return (localStorage.getItem('sunbeth_sticky_header') || 'true') === 'true'; } catch { return true; }
  });
  useEffect(() => {
    // Persist the chosen theme, but avoid forcing the DOM attribute here to prevent
    // racing against ThemeController/TenantProvider initial application.
    try { localStorage.setItem('sunbeth_theme', theme); } catch { /* ignore */ }
  }, [theme]);
  useEffect(() => {
    try { localStorage.setItem('sunbeth_sticky_header', stickyHeader ? 'true' : 'false'); } catch { /* ignore */ }
  }, [stickyHeader]);
  const navigate = useNavigate();
  const location = useLocation();
  const prevAccount = useRef(account);
  const [pending, setPending] = useState<number | null>(null);
  const [dueBy, setDueBy] = useState<string | null>(null);
  const duePoliciesChecked = useRef<boolean>(false);

  useEffect(() => {
    info('Layout mounted');
    const compute = async () => {
      // if not signed in yet, show neutral state
      if ((!account || !token) && !isExternal) { 
        setPending(null); 
        setDueBy(null); 
        return; 
      }

      // Live mode: fetch batches + per-batch progress
      try {
        // Use current token from context
        let list: any[] = [];
        try { 
          const email: string | undefined = (account?.username || externalUser?.email || undefined)?.toLowerCase();
          list = await getBatches(token || undefined, email); 
        } catch { 
          list = []; 
        }
        
        if (!Array.isArray(list) || list.length === 0) { 
          setPending(0); 
          setDueBy(null); 
          return; 
        }
        
        let pendingTotal = 0;
        const incompletes: Array<{ due?: string | null }> = [];
        
        for (const b of list) {
          try {
            const email: string | undefined = (account?.username || externalUser?.email || undefined)?.toLowerCase();
            const p = await getUserProgress(b.toba_batchid, token || undefined, undefined, email);
            const total = p.total ?? 0; 
            const acked = p.acknowledged ?? 0;
            const remain = Math.max(0, total - acked);
            pendingTotal += remain;
            if ((p.percent ?? 0) < 100) incompletes.push({ due: b.toba_duedate });
          } catch { /* ignore */ }
        }
        
        setPending(pendingTotal);
        
        if (incompletes.length) {
          const dates = incompletes.map(i => i.due).filter(Boolean) as string[];
          if (dates.length) {
            const min = dates.reduce((a, d) => (new Date(d) < new Date(a) ? d! : a!));
            setDueBy(min);
          } else setDueBy(null);
        } else setDueBy(null);
      } catch {
        setPending(null);
        setDueBy(null);
      }
    };

    compute();
    const onProgress = () => compute();
    window.addEventListener('sunbeth:progressUpdated', onProgress as EventListener);
    return () => {
      window.removeEventListener('sunbeth:progressUpdated', onProgress as EventListener);
    };
  }, [account, token, isExternal, externalUser?.email]);

  // Proactively prompt for due policies on login
  useEffect(() => {
  const email = (account?.username || externalUser?.email || undefined)?.toLowerCase();
    if (!email || duePoliciesChecked.current) return;
    duePoliciesChecked.current = true;
    (async () => {
      try { await enforceDuePolicies(email); } catch {}
    })();
  }, [account?.username, externalUser?.email]);
  const rbac = useRBAC();

  // If External Support is disabled while an external user is signed in, log them out and route to landing
  useEffect(() => {
    if (flagsLoaded && !externalSupport && isExternal) {
      try { externalLogout(); } catch { /* ignore */ }
      try { navigate('/', { replace: true }); } catch { /* ignore */ }
    }
  }, [flagsLoaded, externalSupport, isExternal, externalLogout, navigate]);
  // Redirect rules around auth transitions for cleaner UX
  useEffect(() => {
    const was = prevAccount.current;
    const now = account;
    // Login occurred
    if (!was && now) {
      // If user is on About (public info) after logging in, send them to Dashboard
      if (location.pathname === '/about') navigate('/', { replace: true });
    }
    // Logout occurred
    if (was && !now) {
      // After logout, ensure we land on the public landing page
      if (location.pathname !== '/') navigate('/', { replace: true });
    }
    prevAccount.current = account;
  }, [account, location.pathname, navigate]);

  // If already authenticated and currently on About, redirect to the dashboard.
  useEffect(() => {
    if (account && location.pathname === '/about') {
      navigate('/', { replace: true });
    }
  }, [account, location.pathname, navigate]);
  const showAside = !!((account || isExternal) && (location.pathname === '/' || location.pathname.startsWith('/dashboard')));
  return (
    <>
      {/* Global busy overlay (dancing logo) */}
      <DancingLogoOverlay />
  <header className={stickyHeader ? 'sticky' : ''}>
        <div className="brand">
          <img src="https://sunbethconcepts.sharepoint.com/:i:/r/sites/CommunicationsandCorporateAffairs/Shared%20Documents/Comms%20Intranet/Logos%20of%20Sunbeth/SGCL%20Coloured%20LOGO.png?csf=1&web=1&e=2IQ9AL" alt="Sunbeth" onError={(e)=>{(e.target as HTMLImageElement).style.opacity = '0.18'; (e.target as HTMLImageElement).alt='Logo'}} />
          <div>
            <div className="h1" style={{ color: '#fff' }}>Sunbeth Document Acknowledgement</div>
            <div className="small" style={{ color: '#fff', opacity: .9 }}>Employee Acknowledgment Portal</div>
            {process.env.NODE_ENV !== 'production' && tenant && (
              <div className="small" style={{ marginTop: 4 }}>
                <span className="badge" title="Active tenant in dev (resolved by X-Tenant-Domain header or host)">Dev Tenant: {tenant.name}{tenant.domain ? ` · ${tenant.domain}` : ''}</span>
              </div>
            )}
          </div>
        </div>

        {/* show auth area when signed-in (MSAL) or as external; else show a light nav */}
        {account ? (
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 12, alignItems: 'center' }}>
            {rbac.isSuperAdmin && (
              <div title="Super Admin (from REACT_APP_SUPER_ADMINS)" style={{ background: '#fee2e2', color: '#991b1b', padding: '6px 8px', borderRadius: 6, fontSize: 13, fontWeight: 700, display: 'flex', alignItems: 'center', gap: 6 }}>
                <span>⚡ Super Admin</span>
              </div>
            )}
            <button className="btn ghost sm" aria-label="Toggle theme" onClick={() => { const next = theme === 'light' ? 'dark' : 'light'; setTheme(next); try { document.documentElement.setAttribute('data-theme', next); window.dispatchEvent(new CustomEvent('sunbeth:themeChanged')); } catch { /* ignore */ } }}>{theme === 'light' ? 'Dark' : 'Light'} Mode</button>
            <button className="btn ghost sm" aria-label="Toggle sticky header" onClick={() => setStickyHeader(s => !s)}>{stickyHeader ? 'Unpin Header' : 'Pin Header'}</button>

            <div style={{ display: 'flex', gap: 12, alignItems: 'center', padding: '6px 8px', background: 'rgba(255,255,255,0.04)', borderRadius: 6 }}>
              <div style={{ width: 36, height: 36, borderRadius: 18, overflow: 'hidden', background: '#fff' }}>
                {photo ? <img src={photo} style={{ width: '100%', height: '100%', objectFit: 'cover' }} alt="avatar" /> : <div style={{ width: '100%', height: '100%', background: '#ccc' }} />}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <div style={{ fontWeight: 700, color: '#fff' }}>{account.name}</div>
                <div style={{ color: '#ddd', fontSize: 13 }}>{account.username}</div>
              </div>
              <button className="btn sm" onClick={() => logout()}>Sign out</button>
            </div>
          </div>
        ) : isExternal ? (
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 12, alignItems: 'center' }}>
            <button className="btn ghost sm" aria-label="Toggle theme" onClick={() => { const next = theme === 'light' ? 'dark' : 'light'; setTheme(next); try { document.documentElement.setAttribute('data-theme', next); window.dispatchEvent(new CustomEvent('sunbeth:themeChanged')); } catch { /* ignore */ } }}>{theme === 'light' ? 'Dark' : 'Light'} Mode</button>
            <button className="btn ghost sm" aria-label="Toggle sticky header" onClick={() => setStickyHeader(s => !s)}>{stickyHeader ? 'Unpin Header' : 'Pin Header'}</button>

            <div style={{ display: 'flex', gap: 12, alignItems: 'center', padding: '6px 8px', background: 'rgba(255,255,255,0.04)', borderRadius: 6 }}>
              <div style={{ width: 36, height: 36, borderRadius: 18, overflow: 'hidden', background: '#e5e7eb', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#111', fontWeight: 700 }}>
                {(externalUser?.name || externalUser?.email || 'U').slice(0,1).toUpperCase()}
              </div>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <div style={{ fontWeight: 700, color: '#fff' }}>{externalUser?.name || 'External User'}</div>
                <div style={{ color: '#ddd', fontSize: 13 }}>{externalUser?.email}</div>
              </div>
              <button className="btn sm" onClick={() => externalLogout()}>Sign out</button>
            </div>
          </div>
        ) : (
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 12, alignItems: 'center' }}>
            <a href="/about" className="small" style={{ color: '#fff', textDecoration: 'none', opacity: .95 }}>About</a>
            <button className="btn ghost sm" aria-label="Toggle theme" onClick={() => { const next = theme === 'light' ? 'dark' : 'light'; setTheme(next); try { document.documentElement.setAttribute('data-theme', next); window.dispatchEvent(new CustomEvent('sunbeth:themeChanged')); } catch { /* ignore */ } }}>{theme === 'light' ? 'Dark' : 'Light'} Mode</button>
            <button className="btn ghost sm" aria-label="Toggle sticky header" onClick={() => setStickyHeader(s => !s)}>{stickyHeader ? 'Unpin Header' : 'Pin Header'}</button>
            <button className="btn sm" onClick={() => login()}>Sign in</button>
          </div>
        )}
      </header>

      <div className={`wrap ${!account ? 'landing-centered' : ''} ${!showAside ? 'centered' : ''}`}>
        <div className="grid">
          <main>
            {children}
          </main>
          {showAside && (
            <aside>
              <div className="card">
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ fontWeight: 700, color: 'var(--primary)' }}>Batch Overview</div>
                    <div className="muted small">Rollout</div>
                  </div>
                  <div style={{ textAlign: 'right' }}>
                    <div style={{ fontWeight: 700, fontSize: 18 }}>{pending ?? '—'}</div>
                    <div className="muted small">You have {pending ?? 0} pending items</div>
                  </div>
                </div>

                <hr style={{ margin: '12px 0', border: 'none', borderTop: '1px solid #f4f4f4' }} />

                <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                  { rbac.canSeeAdmin && <Link to="/admin"><button className="btn full sm">Admin View</button></Link> }
                </div>

                <div style={{ height: 12 }} />
                <div className="muted small">Due by: {dueBy || '—'}</div>
                <div style={{ height: 6 }} />
                <div className="muted small">Assigned to: All staff</div>
              </div>
            </aside>
          )}
        </div>

        <footer>© 2025 Sunbeth Global Concept. All Rights Reserved.</footer>
      </div>
    </>
  );
};

export default Layout;
