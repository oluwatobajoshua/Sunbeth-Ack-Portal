// Lightweight runtime logger that records recent messages to window for UI inspection
type LogLevel = 'debug' | 'info' | 'warn' | 'error';
export const initLogger = () => {
  if (!(window as any).__sunbethLogs) {
    (window as any).__sunbethLogs = [] as Array<{ ts: string; level: LogLevel; msg: string; meta?: any }>;
  }
};

export const log = (level: LogLevel, msg: string, meta?: any) => {
  try {
    const entry = { ts: new Date().toISOString(), level, msg, meta };
    // console output
    if (level === 'error') console.error('[sunbeth]', msg, meta);
    else if (level === 'warn') console.warn('[sunbeth]', msg, meta);
    else console.log('[sunbeth]', msg, meta);
    // store in global buffer (keep last 200)
    const w = window as any;
    if (!w.__sunbethLogs) initLogger();
    w.__sunbethLogs.push(entry);
    if (w.__sunbethLogs.length > 200) w.__sunbethLogs.shift();
  } catch (e) {
    // ignore logging errors
  }
};

export const debug = (msg: string, meta?: any) => log('debug', msg, meta);
export const info = (msg: string, meta?: any) => log('info', msg, meta);
export const warn = (msg: string, meta?: any) => log('warn', msg, meta);
export const error = (msg: string, meta?: any) => log('error', msg, meta);
