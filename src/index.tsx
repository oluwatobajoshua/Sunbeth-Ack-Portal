import React from 'react';
import ReactDOM from 'react-dom/client';
import './styles.css';
import './sunbeth.css';
import App from './App';
import { initLogger, info } from './diagnostics/logger';
import { ErrorBoundary } from './diagnostics/ErrorBoundary';
import { DebugConsole } from './diagnostics/DebugConsole';
import GlobalToast from './diagnostics/GlobalToast';
import { initDiagnostics } from './diagnostics/bootstrap';
import { runAuthAndGraphCheck } from './diagnostics/health';
// Microsoft Graph Toolkit provider setup
import { msalInstance } from './services/msalConfig';

initLogger();
info('index.tsx initializing app');
initDiagnostics();
try { (window as any).__sunbethRunDiagnostics = runAuthAndGraphCheck } catch {}

// Bootstrap asynchronously so we can initialize MSAL v3 before any API calls
(async () => {
  try { await msalInstance.initialize(); } catch { /* ignore */ }

  // Initialize MGT Msal2Provider so components and any MGT usage share auth state
  try {
    const enableMgt = (process.env.REACT_APP_ENABLE_MGT || '').toLowerCase() === 'true';
    const isMock = process.env.REACT_APP_USE_MOCK === 'true';
    const isBrowser = typeof window !== 'undefined';
    if (enableMgt && !isMock && isBrowser) {
      // Only attempt to load MGT dynamically if explicitly enabled and compatible
      const pca: any = (await import('./services/msalConfig')).msalInstance as any;
      const hasGetLogger = typeof pca?.getLogger === 'function';
      if (hasGetLogger) {
        const { Providers } = await import('@microsoft/mgt-element');
        const { Msal2Provider } = await import('@microsoft/mgt-msal2-provider');
        Providers.globalProvider = new Msal2Provider({
          clientId: process.env.REACT_APP_CLIENT_ID as string,
          authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
          redirectUri: (typeof window !== 'undefined' ? window.location.origin : '/'),
          loginType: 'popup',
          scopes: ['User.Read', 'Group.Read.All', 'openid', 'profile']
        } as any);
      } else {
        info('MGT provider skipped: incompatible msal-browser version (no getLogger). Set REACT_APP_ENABLE_MGT=false or upgrade MGT.');
      }
    }
  } catch (e) { /* ignore if MGT is not available */ }

  const root = ReactDOM.createRoot(document.getElementById('root') as HTMLElement);
  root.render(
    <React.StrictMode>
      <ErrorBoundary>
          <App />
          {/* DebugConsole: show only in mock mode */}
          { (process.env.REACT_APP_USE_MOCK === 'true') && <DebugConsole /> }
          <GlobalToast />
      </ErrorBoundary>
    </React.StrictMode>
  );
})();
