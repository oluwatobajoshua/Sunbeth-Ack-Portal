/* eslint-disable max-lines-per-function */
import React, { useEffect, useState } from 'react';
import { getApiBase, isSQLiteEnabled } from '../../utils/runtimeConfig';
import { confirmDialog, showToast } from '../../utils/alerts';
import { createBusiness, deleteBusiness, updateBusiness } from '../../services/dbService';

type Biz = { id: number; name: string; code?: string; isActive?: boolean; description?: string };

const apiBase = () => (getApiBase() as string) || '';
const sqliteOn = () => isSQLiteEnabled() && !!apiBase();

const BusinessesManager: React.FC<{ canEdit: boolean }> = ({ canEdit }) => {
  const [items, setItems] = useState<Biz[]>([]);
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState<{ name: string; code: string; isActive: boolean; description: string }>({ name: '', code: '', isActive: true, description: '' });
  const [editRow, setEditRow] = useState<Record<number, Partial<Biz>>>({});

  const load = async () => {
    if (!sqliteOn()) return;
    try {
      const res = await fetch(`${apiBase()}/api/businesses`);
      const j = await res.json();
      setItems(Array.isArray(j) ? j : []);
    } catch {
      setItems([]);
    }
  };
  useEffect(() => { load(); }, []);

  const create = async () => {
    if (!canEdit || !sqliteOn()) return;
    const name = form.name.trim(); if (!name) { showToast('Enter a business name', 'warning'); return; }
    setBusy(true);
    try {
      await createBusiness({ 
        name, 
        code: form.code || undefined, 
        isActive: !!form.isActive, 
        description: form.description || undefined 
      });
      setForm({ name: '', code: '', isActive: true, description: '' });
      await load();
      showToast('Business created', 'success');
    } catch { showToast('Failed to create business', 'error'); }
    finally { setBusy(false); }
  };

  const save = async (id: number) => {
    if (!canEdit || !sqliteOn()) return;
    const row = editRow[id]; if (!row) return;
    setBusy(true);
    try {
      await updateBusiness(id, row);
      setEditRow(prev => { const p = { ...prev }; delete p[id]; return p; });
      await load();
      showToast('Business updated', 'success');
    } catch { showToast('Failed to update business', 'error'); }
    finally { setBusy(false); }
  };

  const del = async (id: number) => {
    if (!canEdit || !sqliteOn()) return;
    const ok = await confirmDialog('Delete this business?', 'This will unassign it from any recipients.', 'Delete', 'Cancel', { icon: 'warning' as any });
    if (!ok) return;
    setBusy(true);
    try {
      await deleteBusiness(id);
      await load();
      showToast('Business deleted', 'success');
    } catch { showToast('Failed to delete business', 'error'); }
    finally { setBusy(false); }
  };

  if (!sqliteOn()) return <div className="small muted">Enable SQLite to manage businesses.</div>;
  return (
    <div style={{ display: 'grid', gap: 12 }}>
      {/* Create form */}
      <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr 1fr auto', gap: 8, alignItems: 'center' }}>
        <input placeholder="Name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
        <input placeholder="Code (optional)" value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} />
        <input placeholder="Description" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
        <button className="btn sm" onClick={create} disabled={!canEdit || busy}>Add</button>
      </div>
      {/* List */}
      <div style={{ maxHeight: 260, overflowY: 'auto', border: '1px solid #eee', borderRadius: 6 }}>
        {items.length === 0 ? (
          <div className="small muted" style={{ padding: 8 }}>No businesses.</div>
        ) : items.map(b => {
          const isEditing = editRow[b.id] != null;
          return (
            <div key={b.id} style={{ display: 'grid', gridTemplateColumns: '1.2fr 0.8fr 1fr auto auto', gap: 8, alignItems: 'center', padding: 8, borderBottom: '1px solid #f5f5f5' }}>
              {isEditing ? (
                <>
                  <input defaultValue={b.name} onChange={e => setEditRow(prev => ({ ...prev, [b.id]: { ...prev[b.id], name: e.target.value } }))} />
                  <input defaultValue={b.code || ''} onChange={e => setEditRow(prev => ({ ...prev, [b.id]: { ...prev[b.id], code: e.target.value } }))} />
                  <input defaultValue={b.description || ''} onChange={e => setEditRow(prev => ({ ...prev, [b.id]: { ...prev[b.id], description: e.target.value } }))} />
                  <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <input type="checkbox" defaultChecked={!!b.isActive} onChange={e => setEditRow(prev => ({ ...prev, [b.id]: { ...prev[b.id], isActive: e.target.checked } }))} /> Active
                  </label>
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <button className="btn ghost sm" onClick={() => setEditRow(prev => { const p = { ...prev }; delete p[b.id]; return p; })}>Cancel</button>
                    <button className="btn sm" onClick={() => save(b.id)} disabled={busy}>Save</button>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{b.name}</div>
                  <div className="small muted">{b.code || '—'}</div>
                  <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{b.description || '—'}</div>
                  <span className="badge" style={{ background: b.isActive ? '#d4edda' : '#e2e3e5', color: b.isActive ? '#155724' : '#383d41' }}>{b.isActive ? 'Active' : 'Inactive'}</span>
                  <div style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                    <button className="btn ghost sm" onClick={() => setEditRow(prev => ({ ...prev, [b.id]: {} }))} disabled={!canEdit}>Edit</button>
                    <button className="btn ghost sm" onClick={() => del(b.id)} disabled={!canEdit}>Delete</button>
                  </div>
                </>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default BusinessesManager;
