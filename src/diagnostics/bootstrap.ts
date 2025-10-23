import { error, warn, info } from './logger';
import { busyPush, busyPop } from '../utils/busy';

export const initDiagnostics = () => {
  // unhandled promise rejections
  window.addEventListener('unhandledrejection', (ev: any) => {
    try {
      const reason = ev?.reason;
      const code = reason?.errorCode || reason?.code;
      const msg: string = (reason?.message || '').toString();
      // Ignore known benign MSAL redirect noise from MGT provider
      if (code === 'state_not_found' || /state_not_found/i.test(msg)) return;
      error('Unhandled promise rejection', reason || ev);
    } catch (e) {
      // ignore
    }
  });

  window.addEventListener('error', (ev: any) => {
    try {
      error('Window error', { message: ev.message, filename: ev.filename, lineno: ev.lineno, colno: ev.colno, error: ev.error });
    } catch (e) {
      // ignore
    }
  });

  info('Diagnostics initialized');
};

/**
 * Install a global fetch wrapper that pushes a busy overlay while network requests are in-flight.
 * Uses heuristics to derive a helpful label from the URL and HTTP method, unless suppressed.
 */
export const installBusyNetworkTracking = () => {
  try {
    const w = window as any;
    if (w.__sunbethBusyFetchInstalled) return;
    const originalFetch = window.fetch.bind(window);

    const isStaticAsset = (u: string) => /\.(css|js|png|jpg|jpeg|svg|ico|map)(\?|#|$)/i.test(u) || /\/static\//i.test(u);

    const labelFrom = (url: string, method: string) => {
      try {
        const u = url.toLowerCase();
        const m = method.toUpperCase();
        if (isStaticAsset(u)) return '';
        if (u.includes('/api/proxy/graph')) return 'Fetching from SharePoint...';
        if (u.includes('/api/proxy')) return 'Fetching document...';
        if (/\/api\/batches\/[^/]+\/documents/i.test(u) && m === 'GET') return 'Loading documents...';
        if (/\/api\/batches\/[^/]+\/documents/i.test(u) && m === 'POST') return 'Saving documents...';
        if (/\/api\/batches\/[^/]+\/recipients/i.test(u) && m === 'POST') return 'Saving recipients...';
        if (/\/api\/batches\/[^/]+\/completions/i.test(u)) return 'Loading completion summary...';
        if (/\/api\/batches\/?(\?|$)/i.test(u) && m === 'GET') return 'Loading batches...';
        if (/\/api\/batches\/full/i.test(u) && m === 'POST') return 'Creating your batch...';
        if (/\/api\/batches\/[^/]+(\?|$)/i.test(u) && m === 'PUT') return 'Updating batch...';
        if (/\/api\/recipients/i.test(u) && m === 'GET') return 'Loading recipients...';
        if (/\/api\/documents/i.test(u) && m === 'GET') return 'Loading documents...';
        if (/\/api\//i.test(u)) {
          if (m === 'GET') return 'Loading...';
          if (m === 'POST') return 'Saving...';
          if (m === 'PUT' || m === 'PATCH') return 'Updating...';
          if (m === 'DELETE') return 'Deleting...';
        }
        return '';
      } catch { return ''; }
    };

    window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
      const method = (init?.method || (input as any)?.method || 'GET').toString();
      const urlStr = typeof input === 'string' ? input : (input instanceof URL ? input.href : (input as Request).url);
      const silence = Boolean((init as any)?.busySilence);
      const custom = (init as any)?.busyLabel as string | undefined;
      const label = custom ?? labelFrom(urlStr, method);
      if (!silence && label) busyPush(label);
      try {
        const res = await originalFetch(input as any, init as any);
        return res;
      } finally {
        if (!silence && label) busyPop();
      }
    };

    w.__sunbethBusyFetchInstalled = true;
    info('Busy network tracking installed');
  } catch (e) {
    warn('Failed to install busy network tracking', e as any);
  }
};
