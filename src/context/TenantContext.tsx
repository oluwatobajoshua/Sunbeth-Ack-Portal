import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';
import { getApiBase } from '../utils/runtimeConfig';

export type TenantInfo = {
  id: number;
  name: string;
  code: string;
  parentId: number | null;
  isActive: boolean;
  isOwner: boolean;
  domain?: string;
  theme?: ThemeConfig | null;
};

export type ThemeConfig = {
  name?: string;
  darkMode?: boolean;
  logoUrl?: string;
  cssVars?: Record<string, string>; // e.g., { '--primary': '#0055aa' }
  colors?: Partial<{
    primary: string; accent: string; bg: string; bgElevated: string; card: string; muted: string;
  }>;
};

type TenantContextType = {
  tenant: TenantInfo | null;
  loading: boolean;
  error: string | null;
  applyTheme: (theme: ThemeConfig | null | undefined) => void;
};

const TenantContext = createContext<TenantContextType>({ tenant: null, loading: true, error: null, applyTheme: () => {} });

function setCssVar(name: string, value: string) {
  try { document.documentElement.style.setProperty(name, value); } catch { /* no-op in SSR */ }
}

function applyThemeToDom(theme?: ThemeConfig | null) {
  if (!theme) return;
  if (theme.cssVars) {
    Object.entries(theme.cssVars).forEach(([k, v]) => setCssVar(k, String(v)));
  }
  if (theme.colors) {
    if (theme.colors.primary) setCssVar('--primary', theme.colors.primary);
    if (theme.colors.accent) setCssVar('--accent', theme.colors.accent);
    if (theme.colors.bg) setCssVar('--bg', theme.colors.bg);
    if (theme.colors.bgElevated) setCssVar('--bg-elevated', theme.colors.bgElevated);
    if (theme.colors.card) setCssVar('--card', theme.colors.card);
    if (theme.colors.muted) setCssVar('--muted', theme.colors.muted);
  }
  try {
    const mode = theme.darkMode ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', mode);
  } catch { /* ignore */ }
}

export const TenantProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const apiBase = (getApiBase() as string) || '';
  const [tenant, setTenant] = useState<TenantInfo | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let mounted = true;
    (async () => {
      setLoading(true);
      setError(null);
      try {
        // Try new effective theme API first
        let initialTheme: ThemeConfig | null = null;
        try {
          const te = await fetch(`${apiBase}/api/theme/effective`, { cache: 'no-store' });
          if (te.ok) {
            const tj = await te.json();
            // choose variant based on user preference or system
            const pref = (() => {
              try { const p = localStorage.getItem('sunbeth_theme'); if (p === 'light' || p === 'dark') return p; } catch {}
              try { return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'; } catch { return 'light'; }
            })() as 'light'|'dark';
            initialTheme = (tj?.theme?.[pref] || tj?.theme?.light || null) as ThemeConfig | null;
          }
        } catch { /* ignore */ }
        const res = await fetch(`${apiBase}/api/tenant`, { cache: 'no-store' });
        if (!res.ok) throw new Error('tenant_fetch_failed');
        const j = await res.json();
        const t = (j && j.tenant) ? j.tenant as TenantInfo : null;
        if (mounted) {
          setTenant(t);
          applyThemeToDom(initialTheme || t?.theme || null);
        }
      } catch (e: any) {
        if (mounted) setError(e?.message || 'Failed to fetch tenant');
      } finally {
        if (mounted) setLoading(false);
      }
    })();
    return () => { mounted = false; };
  }, [apiBase]);

  const value = useMemo(() => ({ tenant, loading, error, applyTheme: applyThemeToDom }), [tenant, loading, error]);

  return (
    <TenantContext.Provider value={value}>{children}</TenantContext.Provider>
  );
};

export function useTenant() {
  return useContext(TenantContext);
}
