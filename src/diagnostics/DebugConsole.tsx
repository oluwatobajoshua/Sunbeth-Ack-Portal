import React, { useEffect, useState } from 'react';

const readLogs = () => (window as any).__sunbethLogs || [];
const readMsalEvents = () => (window as any).__sunbethMsalEvents || [];

export const DebugConsole: React.FC = () => {
  const [logs, setLogs] = useState<any[]>(readLogs());

  useEffect(() => {
    const iv = setInterval(() => setLogs(readLogs()), 700);
    return () => clearInterval(iv);
  }, []);

  const isDevUser = () => {
    try {
      // Only show when MOCK mode is enabled
      if (process.env.REACT_APP_USE_MOCK === 'true') return true;
    } catch (e) {}
    return false;
  };

  return (
  <div style={{ position: 'fixed', right: 10, bottom: 10, width: 420, maxHeight: 340, overflow: 'auto', background: 'rgba(0,0,0,0.85)', color: '#fff', fontSize: 12, padding: 10, borderRadius: 6, zIndex: 9999, display: isDevUser() ? 'block' : 'none' }}>
      {/* Last Error */}
      <div style={{ marginBottom: 8, padding: 8, background: 'rgba(200,0,0,0.12)', borderRadius: 4 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <strong>Last Error</strong>
          <button onClick={() => { const l = readLogs().filter((x:any)=>x.level==='error').slice(-1)[0]; navigator.clipboard?.writeText(JSON.stringify(l, null, 2)); }}>Copy</button>
        </div>
        <div style={{ fontSize: 12, color: '#ffd1d1' }}>{(() => { const l = readLogs().filter((x:any)=>x.level==='error').slice(-1)[0]; return l ? `${l.ts} ${l.msg}` : 'No errors recorded'; })()}</div>
      </div>
      <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
        <strong>Debug Console</strong>
        <div>
          <label style={{ marginRight: 8 }}>
            <input type="checkbox" checked={(window as any).__sunbethDevUser === true} onChange={(e) => { (window as any).__sunbethDevUser = e.target.checked; setLogs(readLogs()); }} /> Show dev controls
          </label>
          <button onClick={() => { navigator.clipboard?.writeText(JSON.stringify(readLogs(), null, 2)); }} style={{ marginRight: 6 }}>Copy</button>
          <button onClick={() => { (window as any).__sunbethLogs = []; setLogs([]); }}>Clear</button>
        </div>
      </div>
      <div>
        <div style={{ marginBottom: 8 }}>
          <strong>MSAL Events</strong>
          {readMsalEvents().slice().reverse().slice(0,6).map((e:any, idx:number) => (
            <div key={idx} style={{ fontSize: 11, color: '#cfc' }}>{e.ts} — {e.type} {e.payload ? JSON.stringify(e.payload) : ''}</div>
          ))}
        </div>
        {logs.slice().reverse().map((l, i) => (
          <div key={i} style={{ marginBottom: 6, borderBottom: '1px solid rgba(255,255,255,0.06)', paddingBottom: 6 }}>
            <div style={{ color: '#9ad', fontSize: 11 }}>{l.ts} <span style={{ color: '#f7c' }}>[{l.level}]</span></div>
            <div style={{ whiteSpace: 'pre-wrap' }}>{l.msg}</div>
            {l.meta && <pre style={{ color: '#ddd', fontSize: 11 }}>{JSON.stringify(l.meta, null, 2)}</pre>}
          </div>
        ))}
      </div>
    </div>
  );
};
