import React from 'react';
import { useLocation } from 'react-router-dom';
import { getApiBase } from '../utils/runtimeConfig';
import { useTenant } from './TenantContext';

// Map route pathnames to module keys used in theme assignments
function resolveModuleFromPath(pathname: string): string | undefined {
  if (pathname.startsWith('/document') || pathname.startsWith('/batch') || pathname.startsWith('/summary')) return 'docack';
  if (pathname.startsWith('/admin')) return 'admin';
  if (pathname.startsWith('/super-admin')) return 'super-admin';
  if (pathname.startsWith('/modules')) return 'modules';
  return undefined;
}

function getUserPreferredMode(): 'light' | 'dark' {
  try {
    const pref = localStorage.getItem('sunbeth_theme');
    if (pref === 'light' || pref === 'dark') return pref;
  } catch { /* ignore */ }
  try {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  } catch { return 'light'; }
}

export const ThemeController: React.FC = () => {
  const { applyTheme } = useTenant();
  const apiBase = (getApiBase() as string) || '';
  const loc = useLocation();

  React.useEffect(() => {
    let aborted = false;
    const controller = new AbortController();

    const run = async () => {
      const moduleKey = resolveModuleFromPath(loc.pathname);
      const qs = moduleKey ? `?module=${encodeURIComponent(moduleKey)}` : '';
      try {
        const res = await fetch(`${apiBase}/api/theme/effective${qs}`, { cache: 'no-store', signal: controller.signal });
        if (!res.ok) return;
        const data = await res.json();
        const mode = getUserPreferredMode();
        const theme = data?.theme?.[mode] || data?.theme?.light || null;
        if (!aborted) applyTheme(theme || null);
      } catch (_) {
        // ignore fetch errors; don't override existing theme
      }
    };

    run();

    // react to system theme changes if user follows system
    const mql = (() => { try { return window.matchMedia('(prefers-color-scheme: dark)'); } catch { return null as any; } })();
    const handle = () => {
      try {
        const pref = localStorage.getItem('sunbeth_theme');
        if (!pref) {
          // user follows system
          run();
        }
  } catch { /* ignore */ }
    };
  if (mql && typeof mql.addEventListener === 'function') mql.addEventListener('change', handle);

    // React immediately when user toggles theme via UI (custom event)
    const onThemeChanged = () => run();
  try { window.addEventListener('sunbeth:themeChanged', onThemeChanged as EventListener); } catch { /* ignore */ }
    // Cross-tab support: if localStorage changes elsewhere
  const onStorage = (e: StorageEvent) => { if (e.key === 'sunbeth_theme') run(); };
  try { window.addEventListener('storage', onStorage); } catch { /* ignore */ }

    return () => {
      aborted = true;
      controller.abort();
      if (mql && typeof mql.removeEventListener === 'function') mql.removeEventListener('change', handle);
      try { window.removeEventListener('sunbeth:themeChanged', onThemeChanged as EventListener); } catch { /* ignore */ }
      try { window.removeEventListener('storage', onStorage); } catch { /* ignore */ }
    };
  }, [apiBase, loc.pathname, applyTheme]);

  // No visible UI
  return null;
};

export default ThemeController;
