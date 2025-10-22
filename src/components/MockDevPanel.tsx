import React, { useEffect, useState } from 'react';
import { useRuntimeMock } from '../utils/runtimeMock';
import { runAuthAndGraphCheck } from '../diagnostics/health';

const MockDevPanel: React.FC = () => {
  const runtimeMock = useRuntimeMock();
  const [open, setOpen] = useState(false);
  const [state, setState] = useState<Record<string, any>>({});
  const [diag, setDiag] = useState<Array<{ name: string; ok: boolean; detail?: string }>>([]);

  if (!runtimeMock) return null;

  const reload = () => {
    const keys = ['mock_user_acks', 'mock_admin_settings', 'mock_user_groups', 'mock_batches', 'mock_docs_by_batch', 'landing_variant'];
    const obj: Record<string, any> = {};
    for (const k of keys) obj[k] = (() => { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch { return localStorage.getItem(k); } })();
    setState(obj);
  };

  useEffect(() => {
    reload();
    const h = () => reload();
    window.addEventListener('mockAck', h as EventListener);
    window.addEventListener('mockDataChanged', h as EventListener);
    return () => {
      window.removeEventListener('mockAck', h as EventListener);
      window.removeEventListener('mockDataChanged', h as EventListener);
    };
  }, []);

  const setGroups = (kind: 'Admin' | 'Manager' | 'Employee' | 'Clear') => {
    if (kind === 'Admin') localStorage.setItem('mock_user_groups', JSON.stringify(['Sunbeth-Portal-Admins']));
    if (kind === 'Manager') localStorage.setItem('mock_user_groups', JSON.stringify(['Sunbeth-Dept-Managers']));
    if (kind === 'Employee') localStorage.setItem('mock_user_groups', JSON.stringify([]));
    if (kind === 'Clear') localStorage.removeItem('mock_user_groups');
    reload();
    try { window.dispatchEvent(new CustomEvent('sunbeth:roleChange', { detail: { role: kind === 'Admin' ? 'Admin' : kind === 'Manager' ? 'Manager' : 'Employee', mock: true } })); } catch {}
  };

  const seedMock = () => {
    const batches = [
      { toba_batchid: '1', toba_name: 'Q4 2025 — Code of Conduct', toba_startdate: '2025-10-01', toba_duedate: '2025-10-31' },
      { toba_batchid: '2', toba_name: 'Q3 2025 — Health & Safety', toba_startdate: '2025-07-01', toba_duedate: '2025-07-31' }
    ];
    const docs: Record<string, any[]> = {
      '1': [
        { toba_documentid: 'd1', toba_title: 'Code of Conduct', toba_version: 'v1.0' },
        { toba_documentid: 'd2', toba_title: 'Anti-Bribery Policy', toba_version: 'v1.1' }
      ],
      '2': [
        { toba_documentid: 'd3', toba_title: 'Health & Safety Guide', toba_version: 'v2.0' }
      ]
    };
    const acks = { '1': [] };
    localStorage.setItem('mock_batches', JSON.stringify(batches));
    localStorage.setItem('mock_docs_by_batch', JSON.stringify(docs));
    localStorage.setItem('mock_user_acks', JSON.stringify(acks));
    reload();
    try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {}
  };

  const clearMock = () => {
    const keys = ['mock_user_acks', 'mock_admin_settings', 'mock_user_groups', 'mock_batches', 'mock_docs_by_batch', 'landing_variant'];
    for (const k of keys) localStorage.removeItem(k);
    reload();
    try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {}
  };

  return (
    <div style={{ position: 'fixed', left: 12, bottom: 12, zIndex: 9999 }}>
      <button className="btn" onClick={() => setOpen(o => !o)}>{open ? 'Close Dev' : 'Dev Panel'}</button>
      {open && (
        <div style={{ width: 360, maxHeight: 480, overflow: 'auto', marginTop: 8 }} className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div style={{ fontWeight: 700 }}>Dev Panel</div>
            <div className="small muted">Mock mode</div>
          </div>
          <hr />

          <div>
            <div style={{ fontWeight: 700 }}>Mock User Groups</div>
            <div style={{ fontSize: 12, color: '#666', marginBottom: 6 }}>Controls role via localStorage</div>
            <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
              <button className="btn sm" onClick={() => setGroups('Admin')}>Set Admin</button>
              <button className="btn sm" onClick={() => setGroups('Manager')}>Set Manager</button>
              <button className="btn sm" onClick={() => setGroups('Employee')}>Set Employee</button>
              <button className="btn sm ghost" onClick={() => setGroups('Clear')}>Clear Groups</button>
            </div>
            <pre style={{ fontSize: 12, background: '#fafafa', padding: 8, marginTop: 6 }}>{JSON.stringify(state['mock_user_groups'], null, 2)}</pre>
          </div>

          <div style={{ marginTop: 12 }}>
            <div style={{ fontWeight: 700 }}>Seed Mock Data</div>
            <div style={{ display: 'flex', gap: 6, marginTop: 6 }}>
              <button className="btn sm" onClick={seedMock}>Seed</button>
              <button className="btn sm ghost" onClick={clearMock}>Clear</button>
              <button className="btn sm ghost" onClick={reload}>Refresh</button>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 8, marginTop: 8 }}>
              <div>
                <div style={{ fontWeight: 700 }}>mock_batches</div>
                <pre style={{ fontSize: 12, background: '#fafafa', padding: 8 }}>{JSON.stringify(state['mock_batches'], null, 2)}</pre>
              </div>
              <div>
                <div style={{ fontWeight: 700 }}>mock_docs_by_batch</div>
                <pre style={{ fontSize: 12, background: '#fafafa', padding: 8 }}>{JSON.stringify(state['mock_docs_by_batch'], null, 2)}</pre>
              </div>
              <div>
                <div style={{ fontWeight: 700 }}>mock_user_acks</div>
                <pre style={{ fontSize: 12, background: '#fafafa', padding: 8 }}>{JSON.stringify(state['mock_user_acks'], null, 2)}</pre>
              </div>
            </div>
          </div>

          <div style={{ marginTop: 12 }}>
            <div style={{ fontWeight: 700 }}>Landing Variant</div>
            <div style={{ fontSize: 12, color: '#666', marginBottom: 6 }}>Controls hero layout on landing page</div>
            <div style={{ display: 'flex', gap: 6 }}>
              <button className="btn sm" onClick={() => { localStorage.setItem('landing_variant','regular'); reload(); try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {} }}>Regular</button>
              <button className="btn sm" onClick={() => { localStorage.setItem('landing_variant','compact'); reload(); try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {} }}>Compact</button>
              <button className="btn sm ghost" onClick={() => { localStorage.removeItem('landing_variant'); reload(); try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {} }}>Clear</button>
            </div>
            <div className="small muted" style={{ marginTop: 6 }}>Current: <strong>{state['landing_variant'] || 'regular'}</strong></div>
          </div>

          <div style={{ marginTop: 12 }}>
            <div style={{ fontWeight: 700 }}>Diagnostics</div>
            <div style={{ display: 'flex', gap: 6, marginTop: 6 }}>
              <button className="btn sm" onClick={async () => { setDiag([]); const r = await runAuthAndGraphCheck(); setDiag(r); }}>Run Diagnostics</button>
              <button className="btn sm ghost" onClick={() => setDiag([])}>Clear</button>
            </div>
            {diag.length > 0 && (
              <div style={{ marginTop: 8 }}>
                <ul style={{ paddingLeft: 16 }}>
                  {diag.map((d,i) => (<li key={i} style={{ color: d.ok ? 'green' : 'crimson' }}>{d.name}: {d.ok ? 'OK' : 'FAIL'} {d.detail ? ` — ${d.detail}` : ''}</li>))}
                </ul>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default MockDevPanel;
