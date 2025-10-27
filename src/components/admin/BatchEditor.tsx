/* eslint-disable max-lines-per-function, complexity, react-hooks/exhaustive-deps, react-hooks/rules-of-hooks, @typescript-eslint/no-empty-function, @typescript-eslint/no-non-null-assertion, no-empty, max-lines, max-depth, no-useless-escape */
import React from 'react';
import { UserGroupSelector } from './UserGroupSelector';
import LocalLibraryPicker from './LocalLibraryPicker';
import SharePointBrowser from './SharePointBrowser';
import { type GraphUser, getGroupMembers } from '../../services/graphUserService';
import { getGraphToken } from '../../services/authTokens';
import { showToast } from '../../utils/alerts';
import { getApiBase } from '../../utils/runtimeConfig';
import { type SimpleDoc } from './DocumentListEditor';

type Business = { id: number; name: string; code?: string; isActive?: boolean };

export type BatchFormState = {
  name: string;
  startDate: string;
  dueDate: string;
  description: string;
  selectedUsers: GraphUser[];
  selectedGroups: any[]; // GraphGroup (typed as any locally to avoid import cycle)
  selectedDocuments: SimpleDoc[];
  notifyByEmail: boolean;
  notifyByTeams: boolean;
};

export type ImportRow = { name: string; status: 'saved' | 'deduped' | 'failed' };

export interface BatchEditorProps {
  // Permissions/flags
  isSuperAdmin: boolean;
  canUploadDocuments: boolean;
  sqliteEnabled: boolean;

  // Form state
  editingBatchId: string | null;
  batchForm: BatchFormState;
  setBatchForm: React.Dispatch<React.SetStateAction<BatchFormState>>;

  // Selector mode
  useModalSelectors: boolean;
  setUseModalSelectors: (val: boolean) => void;
  modalToggleKey: string;
  usersModalOpen: boolean;
  setUsersModalOpen: (v: boolean) => void;
  docsModalOpen: boolean;
  setDocsModalOpen: (v: boolean) => void;

  // Import progress
  importBusy: boolean;
  importTotal: number;
  importDone: number;
  importRows: ImportRow[];
  setImportBusy: (v: boolean) => void;
  setImportTotal: (v: number) => void;
  setImportDone: (updater: (v: number) => number | number) => void;
  setImportRows: (updater: (rows: ImportRow[]) => ImportRow[] | ImportRow[]) => void;

  // Businesses and mapping
  businesses: Business[];
  mappingUsers: GraphUser[];
  expandGroupsForMapping: () => void;
  businessMap: Record<string, number | null>;
  setUserBusiness: (emailOrUpn: string, businessId: number | null) => void;
  applyBusinessToAll: (businessId: number | null) => void;
  setBusinessMap: (map: Record<string, number | null>) => void;
  defaultBusinessId: number | '';
  setDefaultBusinessId: (v: number | '') => void;

  // Helpers
  mergeDocuments: (prev: SimpleDoc[], incoming: SimpleDoc[]) => SimpleDoc[];
  removeSelectedDoc: (idx: number) => Promise<void> | void;
  saveBatch: () => Promise<void>;
}

const BatchEditor: React.FC<BatchEditorProps> = (props) => {
  const {
    canUploadDocuments,
    sqliteEnabled,
    editingBatchId,
    batchForm,
    setBatchForm,
    useModalSelectors,
    setUseModalSelectors,
    modalToggleKey,
    setUsersModalOpen,
    setDocsModalOpen,
    importBusy,
    importTotal,
    importDone,
    importRows,
    businesses,
    mappingUsers,
    expandGroupsForMapping,
    
    applyBusinessToAll,
    defaultBusinessId,
    setDefaultBusinessId,
    
    saveBatch,
  } = props;

  return (
    <div>
      <h2 style={{ fontSize: 18, marginBottom: 16 }}>{editingBatchId ? 'Edit Batch' : 'Create New Batch'}</h2>

      {/* Batch Details */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 24 }}>
        <div>
          <label className="small" htmlFor="batchName">Batch Name:</label>
          <input
            id="batchName"
            type="text"
            value={batchForm.name}
            onChange={e => setBatchForm({ ...batchForm, name: e.target.value })}
            placeholder="Q1 2025 - Code of Conduct"
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
          />
        </div>
        <div>
          <label className="small" htmlFor="batchDescription">Description:</label>
          <input
            id="batchDescription"
            type="text"
            value={batchForm.description}
            onChange={e => setBatchForm({ ...batchForm, description: e.target.value })}
            placeholder="Annual policy acknowledgement"
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
          />
        </div>
        <div>
          <label className="small" htmlFor="batchStart">Start Date:</label>
          <input
            id="batchStart"
            type="date"
            value={batchForm.startDate}
            onChange={e => setBatchForm({ ...batchForm, startDate: e.target.value })}
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
          />
        </div>
        <div>
          <label className="small" htmlFor="batchDue">Due Date:</label>
          <input
            id="batchDue"
            type="date"
            value={batchForm.dueDate}
            onChange={e => setBatchForm({ ...batchForm, dueDate: e.target.value })}
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
          />
        </div>
      </div>

      {/* Assignment Section */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
        <div className="small muted">Choose how you want to select recipients and documents.</div>
        <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input
            type="checkbox"
            checked={useModalSelectors}
            onChange={e => {
              setUseModalSelectors(e.target.checked);
              try { localStorage.setItem(modalToggleKey, e.target.checked ? 'true' : 'false'); } catch {}
            }}
          />
          Use modal selectors
        </label>
      </div>

      {!useModalSelectors ? (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 24, marginBottom: 24 }}>
          <UserGroupSelector onSelectionChange={(selection: any) => setBatchForm({ ...batchForm, selectedUsers: selection.users, selectedGroups: selection.groups })} />

          {/* Import progress banner (inline) */}
          {importBusy && (
            <div className="small" style={{ background: '#fff8e1', border: '1px solid #ffe0b2', padding: 8, borderRadius: 6 }}>
              Importing to Library... {importDone}/{importTotal}
              <div className="progressBar" aria-hidden="true" style={{ marginTop: 6 }}><i style={{ width: `${importTotal ? Math.round((importDone/importTotal)*100) : 0}%` }} /></div>
              {importRows.length > 0 && (
                <div style={{ marginTop: 6, maxHeight: 120, overflowY: 'auto', display: 'grid', gap: 4 }}>
                  {importRows.map((r, i) => (
                    <div key={i} style={{ display: 'flex', justifyContent: 'space-between' }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.name}</span>
                      <span className="badge" style={{ background: r.status==='failed'?'#f8d7da':(r.status==='deduped'?'#e2e3e5':'#d4edda'), color: '#333' }}>{r.status}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          <LocalLibraryPicker onAdd={(docs) => setBatchForm(prev => ({
            ...prev,
            selectedDocuments: props.mergeDocuments(prev.selectedDocuments, docs)
          }))} />
          <SharePointBrowser canUpload={!!canUploadDocuments} onDocumentSelect={async (spDocs) => {
            // Import SharePoint selections to server library with progress/dedupe status
            try {
              const base = (getApiBase() as string) || '';
              const token = await getGraphToken(['Sites.Read.All','Files.Read.All']);
              props.setImportBusy(true); props.setImportTotal(spDocs.length); props.setImportDone(() => 0); props.setImportRows(() => []);
              const imported: SimpleDoc[] = [];
              let dedupedCount = 0, failed = 0;
              for (const d of spDocs) {
                const driveId = (d as any)?.parentReference?.driveId;
                const itemId = (d as any)?.id;
                const name = d.name;
                if (!base || !driveId || !itemId || !token) {
                  imported.push({ title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint' });
                  props.setImportDone(v => v + 1);
                  props.setImportRows((rows) => [...rows, { name, status: 'failed' }]);
                  failed++;
                  continue;
                }
                try {
                  const res = await fetch(`${base}/api/library/save-graph`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${token}` }, body: JSON.stringify({ driveId, itemId, name }) });
                  const j = await res.json().catch(() => null);
                  const localUrl = j?.url ? `${base}${j.url}` : undefined;
                  // Preserve original SharePoint link in url; store server copy in localUrl
                  const doc: SimpleDoc = { title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint', localFileId: j?.id ?? null, localUrl: localUrl || null };
                  imported.push(doc);
                  props.setImportRows((rows) => [...rows, { name, status: j?.deduped ? 'deduped' : 'saved' }]);
                  if (j?.deduped) dedupedCount++;
                } catch {
                  imported.push({ title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint', localFileId: null, localUrl: null });
                  props.setImportRows((rows) => [...rows, { name, status: 'failed' }]);
                  failed++;
                } finally {
                  props.setImportDone(v => v + 1);
                }
              }
              setBatchForm(prev => ({ ...prev, selectedDocuments: props.mergeDocuments(prev.selectedDocuments, imported) }));
              showToast(`Imported ${imported.length - failed} • deduped ${dedupedCount}${failed ? ` • failed ${failed}` : ''}`, failed ? 'warning' : 'success');
            } catch (e) {
              setBatchForm(prev => ({
                ...prev,
                selectedDocuments: props.mergeDocuments(prev.selectedDocuments, spDocs.map(d => ({ title: d.name, url: d.webUrl, version: 1, requiresSignature: false, driveId: (d as any)?.parentReference?.driveId, itemId: (d as any)?.id, source: 'sharepoint' })))
              }));
            } finally {
              props.setImportBusy(false);
            }
          }} />

          {/* Selected documents list with remove control */}
          <div className="card" style={{ padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 700 }}>Selected Documents</div>
                <div className="small muted">{batchForm.selectedDocuments.length} document(s)</div>
              </div>
            </div>
            {batchForm.selectedDocuments.length === 0 ? (
              <div className="small muted">No documents selected yet. Use Library or SharePoint above.</div>
            ) : (
              <div style={{ maxHeight: 240, overflowY: 'auto', display: 'grid', gap: 6 }}>
                {batchForm.selectedDocuments.map((d, idx) => (
                  <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8, alignItems: 'center' }}>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.title}</div>
                      <div className="small" style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                        {d.source === 'sharepoint' && <span className="badge">sharepoint</span>}
                        {d.localUrl && <span className="badge">local</span>}
                        {d.source === 'sharepoint' && d.localUrl && <span className="badge" title="Server backup created">backed up</span>}
                        {(d.localUrl || d.url) && (
                          <a href={(d.localUrl || d.url)!} target="_blank" rel="noreferrer" className="small">View ↗</a>
                        )}
                      </div>
                    </div>
                    <button className="btn ghost sm" onClick={() => props.removeSelectedDoc(idx)} title="Remove from batch">✕</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Business Mapping */}
          <div className="card" style={{ padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 700 }}>Business Mapping</div>
                <div className="small muted">Assign each selected user to a business</div>
              </div>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
                <span className="small muted">Default:</span>
                <select value={String(defaultBusinessId)} onChange={e => setDefaultBusinessId(e.target.value ? Number(e.target.value) : '')}>
                  <option value="">—</option>
                  {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                </select>
                <span className="small muted">Apply to all:</span>
                <select onChange={e => applyBusinessToAll(e.target.value ? Number(e.target.value) : null)} defaultValue="">
                  <option value="">—</option>
                  {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                </select>
                {batchForm.selectedGroups.length > 0 && (
                  <button className="btn ghost sm" onClick={expandGroupsForMapping} title="Load members of selected groups for per-user mapping">Expand groups for mapping</button>
                )}
                <span className="small muted">Users to map: {mappingUsers.length}</span>
              </div>
            </div>
            {!sqliteEnabled && <div className="small muted">Enable SQLite to load businesses.</div>}
            {sqliteEnabled && businesses.length === 0 && <div className="small muted">No businesses found.</div>}
            {sqliteEnabled && businesses.length > 0 && (
              <div style={{ maxHeight: 260, overflowY: 'auto', display: 'grid', gap: 8 }}>
                {mappingUsers.map(u => {
                  const email = (u.mail || (u as any).userPrincipalName || '').trim().toLowerCase();
                  const sel = props.businessMap[email] ?? '';
                  return (
                    <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, alignItems: 'center' }}>
                      <div>
                        <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</div>
                        <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                      </div>
                      <select value={String(sel)} onChange={e => props.setUserBusiness(email, e.target.value ? Number(e.target.value) : null)}>
                        <option value="">— No business —</option>
                        {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                      </select>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 16, marginBottom: 24 }}>
          {/* Recipients Summary Card */}
          <div className="card" style={{ padding: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontWeight: 700 }}>Assign to Users & Groups</div>
              <div className="small muted">{batchForm.selectedUsers.length} users, {batchForm.selectedGroups.length} groups selected</div>
            </div>
            <button className="btn sm" onClick={() => setUsersModalOpen(true)}>Edit selection</button>
          </div>

          {/* Documents Summary Card */}
          <div className="card" style={{ padding: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <div style={{ fontWeight: 700 }}>Documents</div>
              <div className="small muted">{batchForm.selectedDocuments.length} document(s) selected</div>
            </div>
            <button className="btn sm" onClick={() => setDocsModalOpen(true)}>Choose documents</button>
          </div>

          {/* Business Mapping Summary Card */}
          <div className="card" style={{ padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <div style={{ fontWeight: 700 }}>Business Mapping</div>
                <div className="small muted">Assign selected users to businesses</div>
              </div>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <span className="small muted">Default:</span>
                <select value={String(defaultBusinessId)} onChange={e => setDefaultBusinessId(e.target.value ? Number(e.target.value) : '')}>
                  <option value="">—</option>
                  {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                </select>
                <span className="small muted">Apply to all:</span>
                <select onChange={e => applyBusinessToAll(e.target.value ? Number(e.target.value) : null)} defaultValue="">
                  <option value="">—</option>
                  {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                </select>
              </div>
            </div>
            {sqliteEnabled && businesses.length > 0 ? (
              <div style={{ marginTop: 12, display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, maxHeight: 200, overflowY: 'auto' }}>
                {mappingUsers.map(u => {
                  const email = (u.mail || (u as any).userPrincipalName || '').trim().toLowerCase();
                  const sel = props.businessMap[email] ?? '';
                  return (
                    <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, alignItems: 'center' }}>
                      <div style={{ overflow: 'hidden' }}>
                        <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</div>
                        <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                      </div>
                      <select value={String(sel)} onChange={e => props.setUserBusiness(email, e.target.value ? Number(e.target.value) : null)}>
                        <option value="">— No business —</option>
                        {businesses.map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                      </select>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="small muted" style={{ marginTop: 8 }}>{!sqliteEnabled ? 'Enable SQLite to load businesses.' : 'No businesses found.'}</div>
            )}
          </div>
        </div>
      )}

      {/* Summary & Create */}
      <div style={{ backgroundColor: '#f8f9fa', padding: 16, borderRadius: 8 }}>
        <h3 style={{ margin: '0 0 12px 0', fontSize: 16 }}>Batch Summary</h3>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 16 }}>
          <div>
            <div className="small muted">Assigned Users:</div>
            <div style={{ fontWeight: 'bold' }}>{batchForm.selectedUsers.length + (batchForm.selectedGroups.reduce((acc: number, g: any) => acc + (g.memberCount || 0), 0))}</div>
          </div>
          <div>
            <div className="small muted">Documents:</div>
            <div style={{ fontWeight: 'bold' }}>{batchForm.selectedDocuments.length}</div>
          </div>
          <div>
            <div className="small muted">Duration:</div>
            <div style={{ fontWeight: 'bold' }}>
              {batchForm.startDate && batchForm.dueDate ?
                Math.ceil((new Date(batchForm.dueDate).getTime() - new Date(batchForm.startDate).getTime()) / (1000 * 60 * 60 * 24)) + ' days'
                : 'Not set'}
            </div>
          </div>
        </div>

        {/* Notification options */}
        <div style={{ display: 'flex', gap: 16, marginTop: 16, alignItems: 'center', flexWrap: 'wrap' }}>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <input type="checkbox" checked={batchForm.notifyByEmail} onChange={e => setBatchForm({ ...batchForm, notifyByEmail: e.target.checked })} />
            <span className="small">Email notification (Microsoft Graph)</span>
          </label>
          <label style={{ display: 'flex', alignItems: 'center', gap: 8, opacity: .6 }} title="Requires Teams Chat.ReadWrite; coming soon">
            <input type="checkbox" checked={batchForm.notifyByTeams} onChange={e => setBatchForm({ ...batchForm, notifyByTeams: e.target.checked })} disabled />
            <span className="small">Teams message (optional)</span>
          </label>
        </div>

        <div style={{ marginTop: 16, display: 'flex', gap: 8 }}>
          <button
            className="btn"
            onClick={saveBatch}
            disabled={!batchForm.name || !batchForm.startDate || !batchForm.dueDate || batchForm.selectedDocuments.length === 0 || (batchForm.selectedUsers.length === 0 && batchForm.selectedGroups.length === 0)}
          >
            {editingBatchId ? 'Save Changes' : 'Create Batch'}
          </button>
          <button className="btn ghost" onClick={() => { setBatchForm({ name: '', startDate: '', dueDate: '', description: '', selectedUsers: [], selectedGroups: [], selectedDocuments: [], notifyByEmail: true, notifyByTeams: false }); props.setBusinessMap({}); setDefaultBusinessId(''); }}>
            {editingBatchId ? 'Cancel Edit' : 'Reset Form'}
          </button>
          <button className="btn ghost" title="Preview expanded recipients" onClick={async () => {
            try {
              const recipientSet = new Set<string>();
              for (const u of batchForm.selectedUsers) {
                const addr = (u.mail || (u as any).userPrincipalName || '').trim();
                if (addr) recipientSet.add(addr.toLowerCase());
              }
              if (batchForm.selectedGroups.length > 0) {
                const token = await getGraphToken(['Group.Read.All','User.Read']);
                const arrays = await Promise.all(batchForm.selectedGroups.map(g => getGroupMembers(token, (g as any).id).catch(() => [])));
                const members = ([] as GraphUser[]).concat(...arrays);
                for (const m of members) {
                  const addr = (m.mail || (m as any).userPrincipalName || '').trim();
                  if (addr) recipientSet.add(addr.toLowerCase());
                }
              }
              const count = recipientSet.size;
              showToast(`Recipient preview: ${count} unique addresses`, 'info');
            } catch (e) {
              showToast('Failed to preview recipients', 'error');
            }
          }}>Preview Recipients</button>
          <button className="btn ghost" title="Grant Graph permissions" onClick={async () => {
            try {
              // Trigger consent prompts for common scopes used in Admin
              await getGraphToken(['User.Read.All','Group.Read.All']);
              await getGraphToken(['Mail.Send']);
              await getGraphToken(['Sites.Read.All','Files.ReadWrite.All']);
              showToast('Permissions granted (if consented)', 'success');
            } catch (e) {
              showToast('Permission grant failed', 'error');
            }
          }}>Grant Permissions</button>
        </div>
      </div>
    </div>
  );
};

export default BatchEditor;
