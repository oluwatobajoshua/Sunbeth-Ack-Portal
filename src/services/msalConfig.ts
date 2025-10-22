import { PublicClientApplication, AccountInfo, LogLevel } from '@azure/msal-browser';
import { info, warn, error as logError } from '../diagnostics/logger';

// Determine if MSAL can be safely used in the current runtime. In Jest/jsdom or
// other non-browser environments, the Web Crypto API may be missing which causes
// MSAL to throw on construction. Also disable MSAL entirely when running in mock mode.
const canUseMsal = (): boolean => {
  try {
    if (process.env.REACT_APP_USE_MOCK === 'true') return false;
    if (typeof window === 'undefined') return false;
    const anyWin: any = window as any;
    return !!(anyWin.crypto && (anyWin.crypto.subtle || anyWin.msCrypto));
  } catch {
    return false;
  }
};

// Create a minimal safe fallback object when MSAL is disabled. Methods either no-op or
// return sensible defaults used by our AuthContext guards. All calls in mock mode avoid MSAL anyway.
const createMsalFallback = () => {
  const noop = () => undefined as any;
  return {
    addEventCallback: noop,
    handleRedirectPromise: async () => null,
    setActiveAccount: noop,
    getAllAccounts: () => [] as any[],
    loginPopup: async () => { throw new Error('MSAL disabled in this environment'); },
    loginRedirect: async () => { throw new Error('MSAL disabled in this environment'); },
    acquireTokenSilent: async () => { throw new Error('MSAL disabled in this environment'); },
    acquireTokenPopup: async () => { throw new Error('MSAL disabled in this environment'); },
    logoutPopup: async () => { /* no-op */ }
  } as unknown as PublicClientApplication;
};

export const msalInstance: PublicClientApplication = canUseMsal()
  ? new PublicClientApplication({
      auth: {
        clientId: process.env.REACT_APP_CLIENT_ID as string,
        authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
        // Use explicit origin so the registered redirect URIs match exactly
        redirectUri: (typeof window !== 'undefined' ? window.location.origin : '/'),
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false },
      system: {
        loggerOptions: {
          loggerCallback: (level, message, containsPii) => {
            try {
              if (containsPii) return;
              const isMock = process.env.REACT_APP_USE_MOCK === 'true';
              if (level === LogLevel.Error) logError('msal', message);
              else if (level === LogLevel.Warning) warn('msal', message);
              else if (isMock) info('msal', message); // suppress info/verbose in live
            } catch (e) { /* ignore */ }
          },
          piiLoggingEnabled: false,
          // Reduce noise in live; keep detailed logs in mock
          logLevel: (process.env.REACT_APP_USE_MOCK === 'true') ? LogLevel.Info : LogLevel.Warning
        }
      }
    })
  : createMsalFallback();

export type { AccountInfo };

// attach an event callback to capture MSAL lifecycle events into our logger
try {
  msalInstance.addEventCallback((ev) => {
    try {
      // Log useful event types
      const type = (ev && (ev as any).eventType) || 'msal_event';
      const payload = (ev && (ev as any).payload) || {};
      const isMock = process.env.REACT_APP_USE_MOCK === 'true';
      if (isMock) info('msal:event', { type, payload });
      try {
        // Only store MSAL events for on-screen diagnostics in mock/local use
        if (isMock) {
          const arr = (window as any).__sunbethMsalEvents = (window as any).__sunbethMsalEvents || [];
          arr.push({ ts: new Date().toISOString(), type, payload });
          if (arr.length > 200) arr.shift();
        }
      } catch (e) { }
    } catch (e) {
      try { logError('msal:event logging failed', e); } catch { }
    }
  });
} catch (e) { /* ignore in environments where msal isn't available yet */ }
