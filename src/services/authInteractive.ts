// Controls whether interactive authentication uses popup or redirect.
// Priority: localStorage override > env var REACT_APP_MSAL_INTERACTIVE_FALLBACK > default 'popup'.

export type InteractiveMode = 'popup' | 'redirect';

const STORAGE_KEY = 'sunbeth:interactiveMode';

export function getInteractiveMode(): InteractiveMode {
  try {
    const ls = localStorage.getItem(STORAGE_KEY) as InteractiveMode | null;
    if (ls === 'popup' || ls === 'redirect') return ls;
  } catch {}
  const env = (process.env.REACT_APP_MSAL_INTERACTIVE_FALLBACK || '').toLowerCase();
  if (env === 'redirect' || env === 'popup') return env as InteractiveMode;
  return 'popup';
}

export function setInteractiveMode(mode: InteractiveMode) {
  try { localStorage.setItem(STORAGE_KEY, mode); } catch {}
}
