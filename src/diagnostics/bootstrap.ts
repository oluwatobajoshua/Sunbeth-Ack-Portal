import { error, warn, info } from './logger';

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
