/* eslint-disable max-lines-per-function */
/* eslint-disable complexity */
import React, { useEffect, useState } from 'react';
import Modal from '../Modal';
import { getApiBase, isSQLiteEnabled } from '../../utils/runtimeConfig';
import { confirmDialog, showToast } from '../../utils/alerts';

type BatchRow = { toba_batchid: string; toba_name: string; toba_startdate?: string; toba_duedate?: string; toba_status?: string };

const apiBase = () => (getApiBase() as string) || '';
const sqliteOn = () => isSQLiteEnabled() && !!apiBase();

const ManageBatches: React.FC<{ canEdit: boolean; onEdit: (id: string) => void; onClone: (id: string) => void }>
  = ({ canEdit, onEdit, onClone }) => {
  const [items, setItems] = useState<Array<BatchRow>>([]);
  const [busy, setBusy] = useState(false);
  const [editing, setEditing] = useState<Record<string, { name: string; startDate: string; dueDate: string; status: string; description: string }>>({});
  const [recOpen, setRecOpen] = useState<{ open: boolean; forBatch?: string; rows: any[] }>({ open: false, rows: [] });

  const load = async () => {
    if (!sqliteOn()) return;
    try {
      const res = await fetch(`${apiBase()}/api/batches`);
      const j = await res.json();
      setItems(Array.isArray(j) ? j : []);
    } catch { setItems([]); }
  };
  useEffect(() => { load(); }, []);

  const del = async (id: string) => {
    if (!canEdit || !sqliteOn()) return;
    const ok = await confirmDialog('Delete this batch?', 'This will remove its documents, recipients, and acknowledgements.', 'Delete', 'Cancel', { icon: 'warning' as any });
    if (!ok) return;
    setBusy(true);
    try {
      const res = await fetch(`${apiBase()}/api/batches/${encodeURIComponent(id)}`, { method: 'DELETE' });
      if (!res.ok) throw new Error('delete_failed');
      await load();
      showToast('Batch deleted', 'success');
    } catch {
      showToast('Failed to delete batch', 'error');
    } finally { setBusy(false); }
  };

  const openRecipients = async (id: string) => {
    try {
      const res = await fetch(`${apiBase()}/api/recipients`);
      const j = await res.json();
      const rows = (Array.isArray(j) ? j : []).filter((r: any) => String(r.batchId) === String(id));
      setRecOpen({ open: true, forBatch: id, rows });
    } catch { setRecOpen({ open: true, forBatch: id, rows: [] }); }
  };

  const save = async (id: string) => {
    const row = editing[id]; if (!row) return;
    setBusy(true);
    try {
      const payload = {
        name: row.name,
        startDate: row.startDate || null,
        dueDate: row.dueDate || null,
        status: row.status ? Number(row.status) : 1,
        description: row.description || null
      };
      const res = await fetch(`${apiBase()}/api/batches/${encodeURIComponent(id)}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!res.ok) throw new Error('update_failed');
      setEditing(prev => { const p = { ...prev }; delete p[id]; return p; });
      await load();
      showToast('Batch updated', 'success');
    } catch {
      showToast('Failed to update batch', 'error');
    } finally { setBusy(false); }
  };

  if (!sqliteOn()) return <div className="small muted">Enable SQLite to manage batches.</div>;
  return (
    <>
      <div style={{ maxHeight: 420, overflowY: 'auto', border: '1px solid #eee', borderRadius: 6 }}>
        {items.length === 0 ? (
          <div className="small muted" style={{ padding: 8 }}>No batches.</div>
        ) : items.map(b => {
          const row = editing[b.toba_batchid];
          const isEditing = !!row;
          return (
            <div key={b.toba_batchid} style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.9fr 0.9fr 0.7fr 1.4fr auto', gap: 8, alignItems: 'center', padding: 8, borderBottom: '1px solid #f5f5f5' }}>
              {isEditing ? (
                <>
                  <input defaultValue={b.toba_name} onChange={e => setEditing(prev => ({ ...prev, [b.toba_batchid]: { ...(prev[b.toba_batchid] || {}), name: e.target.value } }))} />
                  <input type="date" defaultValue={b.toba_startdate || ''} onChange={e => setEditing(prev => ({ ...prev, [b.toba_batchid]: { ...(prev[b.toba_batchid] || {}), startDate: e.target.value } }))} />
                  <input type="date" defaultValue={b.toba_duedate || ''} onChange={e => setEditing(prev => ({ ...prev, [b.toba_batchid]: { ...(prev[b.toba_batchid] || {}), dueDate: e.target.value } }))} />
                  <select defaultValue={b.toba_status || '1'} onChange={e => setEditing(prev => ({ ...prev, [b.toba_batchid]: { ...(prev[b.toba_batchid] || {}), status: e.target.value } }))}>
                    <option value="1">Active</option>
                    <option value="0">Inactive</option>
                  </select>
                  <input placeholder="Description" onChange={e => setEditing(prev => ({ ...prev, [b.toba_batchid]: { ...(prev[b.toba_batchid] || {}), description: e.target.value } }))} />
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <button className="btn ghost sm" onClick={() => setEditing(prev => { const p = { ...prev }; delete p[b.toba_batchid]; return p; })}>Cancel</button>
                    <button className="btn sm" onClick={() => save(b.toba_batchid)} disabled={!canEdit || busy}>Save</button>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{b.toba_name}</div>
                  <div className="small muted">{b.toba_startdate || '—'}</div>
                  <div className="small muted">{b.toba_duedate || '—'}</div>
                  <span className="badge" style={{ background: (b.toba_status || '1') === '1' ? '#d4edda' : '#e2e3e5', color: (b.toba_status || '1') === '1' ? '#155724' : '#383d41' }}>{(b.toba_status || '1') === '1' ? 'Active' : 'Inactive'}</span>
                  <div className="small muted" />
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <a href={`/batch/${b.toba_batchid}`}><button className="btn ghost sm">View</button></a>
                    <button className="btn ghost sm" onClick={() => openRecipients(b.toba_batchid)}>Recipients</button>
                    <button className="btn ghost sm" onClick={() => onEdit(b.toba_batchid)} disabled={!canEdit}>Edit</button>
                    <button className="btn ghost sm" onClick={() => onClone(b.toba_batchid)} disabled={!canEdit}>Clone</button>
                    <button className="btn ghost sm" onClick={() => del(b.toba_batchid)} disabled={!canEdit || busy}>Delete</button>
                  </div>
                </>
              )}
            </div>
          );
        })}
      </div>
      {/* Recipients Modal */}
      <Modal open={recOpen.open} onClose={() => setRecOpen({ open: false, rows: [] })} title={`Recipients for Batch ${recOpen.forBatch || ''}`} width={700}>
        {recOpen.rows.length === 0 ? (
          <div className="small muted">No recipients found.</div>
        ) : (
          <div style={{ maxHeight: 360, overflowY: 'auto', display: 'grid', gap: 8 }}>
            {recOpen.rows.map((r: any, i: number) => (
              <div key={i} style={{ display: 'grid', gridTemplateColumns: '1.6fr 1fr 1fr 1fr', gap: 8 }}>
                <div>
                  <div style={{ fontWeight: 500 }}>{r.displayName || r.email}</div>
                  <div className="small muted">{r.email}</div>
                </div>
                <div className="small muted">{r.department || '—'}</div>
                <div className="small muted">{r.jobTitle || '—'}</div>
                <div className="small muted">{r.primaryGroup || '—'}</div>
              </div>
            ))}
          </div>
        )}
      </Modal>
    </>
  );
};

export default ManageBatches;
