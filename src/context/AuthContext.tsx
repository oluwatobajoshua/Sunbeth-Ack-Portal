/**
 * AuthContext: Centralizes authentication state and token acquisition (MSAL).
 *
 * Behavior:
 * - Processes redirect responses on mount and only attaches an account after acquiring a token.
 * - Popup-first with redirect fallback to avoid popup blockers.
 * - Suppresses auto re-attach immediately after logout via a suppression flag.
 *
 * Exposed API:
 * - account, token, photo
 * - login(), logout(), getToken(scopes)
 */
import React, { createContext, useContext, useEffect, useState } from 'react';
import { msalInstance, AccountInfo } from '../services/msalConfig';
import { info, warn, error as logError } from '../diagnostics/logger';
import { getInteractiveMode } from '../services/authInteractive';

type AuthCtx = {
  account: AccountInfo | null;
  token: string | null;
  photo?: string | null;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  getToken: (scopes?: string[]) => Promise<string | null>;
};

const defaultCtx: AuthCtx = {
  account: null,
  token: null,
  photo: null,
  login: async () => { },
  logout: async () => { },
  getToken: async () => null
};

export const AuthContext = createContext<AuthCtx>(defaultCtx);
export const useAuth = () => useContext(AuthContext);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  // AuthProvider is mounted above BrowserRouter, so use window.location for navigation
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [photo, setPhoto] = useState<string | null>(null);
  const SUPPRESS_KEY = 'sunbeth:suppressAutoLogin';
  const [loggingIn, setLoggingIn] = useState(false);

  useEffect(() => {
    // Always process redirect responses first so loginRedirect can complete
    try {
      const hasAuthParams = (() => {
        try {
          const h = window.location.hash || '';
          const s = window.location.search || '';
          return /[?#&](code|error|id_token)=/.test(h) || /[?#&](code|error|id_token)=/.test(s);
        } catch { return false; }
      })();
      (hasAuthParams ? msalInstance.handleRedirectPromise() : Promise.resolve(null)).then((res) => {
        if (res && res.account) {
          info('AuthContext handleRedirectPromise: account present', { username: res.account.username });
          try { localStorage.removeItem(SUPPRESS_KEY); } catch {}
          msalInstance.setActiveAccount(res.account);
          // Only attach account after acquiring a token
          msalInstance
            .acquireTokenSilent({ scopes: ['User.Read'], account: res.account })
            .then(r => { setToken(r.accessToken); setAccount(res.account!); info('AuthContext redirect: token acquired, account attached'); })
            .catch(async err => {
              warn('AuthContext redirect: token acquisition failed (silent). Attempting popup fallback.', err);
              try {
                const pr = await msalInstance.acquireTokenPopup({ scopes: ['User.Read'] });
                setToken(pr.accessToken);
                setAccount(res.account!);
                info('AuthContext redirect: popup token acquired, account attached');
              } catch (ee) {
                logError('AuthContext redirect: popup token acquisition failed', ee);
              }
            });
        }
      }).catch(err => {
        warn('handleRedirectPromise error', err);
      });
    } catch (e) { /* ignore */ }

    // If auto-login suppression is active (e.g., just logged out), don't auto attach account from cache
    let suppressed = false;
    try { suppressed = localStorage.getItem(SUPPRESS_KEY) === 'true'; } catch {}
    if (suppressed) {
      info('AuthContext: auto-login suppressed (skipping cached account attach)');
      return;
    }

    info('AuthContext checking MSAL accounts');
    const accts = msalInstance.getAllAccounts();
    if (accts.length > 0) {
      info('AuthContext found existing account', { count: accts.length });
      // Only attach cached account if token can be acquired silently
      msalInstance
        .acquireTokenSilent({ scopes: ['User.Read'], account: accts[0] })
        .then(r => { setToken(r.accessToken); setAccount(accts[0]); info('AuthContext attached cached account with token'); })
        .catch(() => { info('AuthContext: cached account exists but token acquisition failed; awaiting manual login'); });
    }
  }, []);

  useEffect(() => {
    if (!account) return;
    info('AuthContext acquiring token silently', { account: account.username });
    msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account })
      .then(r => { info('AuthContext acquired token silently'); setToken(r.accessToken); })
      .catch(async (e) => {
        warn('acquireTokenSilent failed, falling back to popup', e);
        try {
          const r = await msalInstance.acquireTokenPopup({ scopes: ['User.Read'] });
          setToken(r.accessToken);
          info('AuthContext acquired token via popup');
        } catch (ee) {
          logError('acquireTokenPopup failed', ee);
        }
      });
  }, [account]);

  // fetch profile photo from Microsoft Graph when token becomes available
  useEffect(() => {
    if (!token) return;
    let active = true;
    let objectUrl: string | null = null;
    (async () => {
      try {
        const res = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!res.ok) {
          setPhoto(null);
          return;
        }
        const blob = await res.blob();
        objectUrl = URL.createObjectURL(blob);
        if (active) setPhoto(objectUrl);
      } catch (e) {
        if (active) setPhoto(null);
      }
    })();

    return () => {
      active = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    };
  }, [token]);

  const login = async () => {
    try {
      if (loggingIn) {
        info('AuthContext: login ignored (already in progress)');
        return;
      }
      setLoggingIn(true);
      info('AuthContext: manual login requested');
      try { localStorage.removeItem(SUPPRESS_KEY); } catch {}
      
      const mode = getInteractiveMode();
      const scopes = ['User.Read', 'openid', 'profile'];
      const res = mode === 'popup'
        ? await msalInstance.loginPopup({ scopes, prompt: 'select_account' })
        : (await msalInstance.loginRedirect({ scopes, prompt: 'select_account' }), null as any);
      info('AuthContext loginPopup result', res);

      const accounts = msalInstance.getAllAccounts();
      info('AuthContext accounts after popup', { count: accounts.length });
      const acct = (accounts && accounts.length) ? accounts[0] : res.account;
      if (!acct) {
        logError('login: no account returned');
        try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Sign in failed (no account returned)' } })); } catch (e) { }
        return;
      }
      msalInstance.setActiveAccount(acct);
      // Acquire token before attaching account
      let access: string | null = null;
      try {
        const tokenResp = await msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account: acct });
        access = tokenResp.accessToken;
        info('AuthContext acquired token silently after login');
      } catch (e) {
        warn('acquireTokenSilent after login failed, attempting popup', e);
        try {
          const mode2 = getInteractiveMode();
          if (mode2 === 'popup') {
            const tr = await msalInstance.acquireTokenPopup({ scopes: ['User.Read'] });
            access = tr.accessToken;
          } else {
            await msalInstance.acquireTokenRedirect({ scopes: ['User.Read'] });
            // redirect will navigate; stop here
            return;
          }
        } catch (ee) {
          logError('acquireTokenPopup after login failed', ee);
        }
      }
      if (!access) {
        try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Sign in failed: could not acquire token' } })); } catch (e) { }
        return;
      }
  setToken(access);
      setAccount(acct);
      try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Signed in' } })); } catch (e) { }

  // No full reload; view will switch to Dashboard based on account state
    } catch (e) {
      logError('login failed (popup)', e);
      // fallback to redirect if popup fails (popup blocked or not supported)
      try {
        info('AuthContext: falling back to loginRedirect');
        await msalInstance.loginRedirect({ scopes: ['User.Read', 'openid', 'profile'] });
        return;
      } catch (er) {
        logError('loginRedirect also failed', er);
        try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Sign in failed' } })); } catch (err) { }
      }
    } finally { setLoggingIn(false); }
  };

  const logout = async () => {
    try {
      info('AuthContext: logout requested');
      await msalInstance.logoutPopup();
      setAccount(null);
      setToken(null);
      setPhoto(null);
      try { localStorage.setItem(SUPPRESS_KEY, 'true'); } catch {}
      try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Signed out' } })); } catch (e) { }
      // No full reload; routes will render Landing based on account=null
    } catch (e) {
      logError('logout failed', e);
      try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Sign out failed' } })); } catch (er) { }
    }
  };

  const getToken = async (scopes: string[] = ['User.Read']) => {
    info('AuthContext getToken', { scopes });
    const acct = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0] || account;
    if (!acct) {
      throw new Error('No active account');
    }
    try {
      const resp = await msalInstance.acquireTokenSilent({ account: acct, scopes });
      return resp.accessToken;
    } catch (e) {
      const mode = getInteractiveMode();
      warn('getToken: acquireTokenSilent failed, using interactive fallback', { mode, error: e as any });
      if (mode === 'popup') {
        // Brief delay helps if another interaction just finished
        await new Promise(r => setTimeout(r, 200));
        try {
          const resp = await msalInstance.acquireTokenPopup({ scopes });
          return resp.accessToken;
        } catch (err) {
          await new Promise(r => setTimeout(r, 400));
          const resp2 = await msalInstance.acquireTokenPopup({ scopes });
          return resp2.accessToken;
        }
      } else {
        await msalInstance.acquireTokenRedirect({ scopes });
        // Redirect initiated; throw to stop current flow
        throw new Error('Interactive redirect initiated for token');
      }
    }
  };

  return <AuthContext.Provider value={{ account, token, photo, login, logout, getToken }}>{children}</AuthContext.Provider>;
};
