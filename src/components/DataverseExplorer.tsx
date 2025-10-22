import React, { useEffect, useMemo, useState } from 'react';
import { getDataverseToken } from '../services/authTokens';
import { DV_SETS, DV_ATTRS } from '../services/dataverseConfig';
import { listRecords, getEntityLogicalNameBySet, getEntityAttributes, createRecord, updateRecord, deleteRecord } from '../services/dataverseService';

const pretty = (v: any) => {
  try { return JSON.stringify(v, null, 2); } catch { return String(v); }
};

const parseJson = (t: string) => {
  try { return JSON.parse(t); } catch { return null; }
};

const guidLike = (s: string) => /[0-9a-fA-F-]{36}/.test(s);

const SetOption: React.FC<{ name: string; value: string }> = ({ name, value }) => (
  <option value={value}>{name} ({value})</option>
);

const DataverseExplorer: React.FC = () => {
  const [selectedSet, setSelectedSet] = useState<string>(DV_SETS.batchesSet);
  const [loading, setLoading] = useState(false);
  const [rows, setRows] = useState<any[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [logicalName, setLogicalName] = useState<string | undefined>(undefined);
  const [attributes, setAttributes] = useState<string[]>([]);
  const [top, setTop] = useState<number>(50);

  // CRUD state
  const [createBody, setCreateBody] = useState<string>('{\n  "toba_name": "Sample"\n}');
  const [updateId, setUpdateId] = useState<string>('');
  const [updateBody, setUpdateBody] = useState<string>('{}');
  const [deleteId, setDeleteId] = useState<string>('');
  const [opResult, setOpResult] = useState<string>('');

  const sets = useMemo(() => ([
    { name: 'Batches', set: DV_SETS.batchesSet },
    { name: 'Documents', set: DV_SETS.documentsSet },
    { name: 'Batch Recipients', set: DV_SETS.batchRecipientsSet },
    { name: 'User Acknowledgements', set: DV_SETS.userAcksSet },
    { name: 'User Progresses', set: DV_SETS.userProgressesSet },
    { name: 'Businesses', set: DV_SETS.businessesSet }
  ]), []);

  const load = async () => {
    if (!(process.env.REACT_APP_ENABLE_DATAVERSE === 'true') || !process.env.REACT_APP_DATAVERSE_URL) {
      setError('Dataverse disabled or URL not set');
      return;
    }
    setLoading(true); setError(null);
    try {
      const t = await getDataverseToken();
      const [logic, rowsData] = await Promise.all([
        getEntityLogicalNameBySet(selectedSet, t).catch(() => undefined),
        listRecords(selectedSet, t, { top })
      ]);
      setRows(Array.isArray(rowsData) ? rowsData : []);
      setLogicalName(logic);
      if (logic) {
        const attrs = await getEntityAttributes(logic, t).catch(() => []);
        setAttributes(attrs);
      } else {
        setAttributes([]);
      }
    } catch (e: any) {
      setError(typeof e?.message === 'string' ? e.message : 'Failed to load records');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, [selectedSet, top]);

  const idField = useMemo(() => logicalName ? `${logicalName}id` : undefined, [logicalName]);
  const appColumnsForSet = (setName: string): string[] => {
    const ensure = (...names: Array<string | undefined>) => names.filter(Boolean) as string[];
    if (setName === DV_SETS.batchesSet) {
      return ensure('toba_batchid', 'toba_name', 'toba_startdate', 'toba_duedate', 'toba_status');
    }
    if (setName === DV_SETS.documentsSet) {
      return ensure('toba_documentid',
        DV_ATTRS.docTitleField || 'toba_title',
        DV_ATTRS.docVersionField || 'toba_version',
        DV_ATTRS.docRequiresSigField || 'toba_requiressignature',
        DV_ATTRS.docUrlField || 'toba_fileurl',
        '_toba_batch_value');
    }
    if (setName === DV_SETS.businessesSet) {
      return ensure('toba_businessid', 'toba_name', 'toba_code', 'toba_isactive');
    }
    if (setName === DV_SETS.batchRecipientsSet) {
      return ensure('toba_batchrecipientid', 'toba_name',
        DV_ATTRS.batchRecipientUserField || 'toba_User',
        DV_ATTRS.batchRecipientEmailField || 'toba_Email',
        DV_ATTRS.batchRecipientDisplayNameField || 'toba_DisplayName',
        DV_ATTRS.batchRecipientDepartmentField || 'toba_Department',
        DV_ATTRS.batchRecipientJobTitleField || 'toba_JobTitle',
        DV_ATTRS.batchRecipientLocationField || 'toba_Location',
        DV_ATTRS.batchRecipientPrimaryGroupField || 'toba_PrimaryGroup',
        '_toba_batch_value', '_toba_business_value');
    }
    if (setName === DV_SETS.userAcksSet) {
      return ensure('toba_useracknowledgementid', 'toba_acknowledged', 'toba_ackdate', DV_ATTRS.ackUserField || 'toba_User', '_toba_batch_value', '_toba_document_value');
    }
    if (setName === DV_SETS.userProgressesSet) {
      return ensure('toba_batchuserprogressid', 'toba_acknowledged', 'toba_totaldocs', DV_ATTRS.ackUserField || 'toba_User', '_toba_batch_value');
    }
    // Fallback: id + name-ish columns if known
    return ensure(idField, 'toba_name');
  };
  const columns = useMemo(() => {
    const preferred = appColumnsForSet(selectedSet);
    if (rows.length === 0) return preferred; // show headers even if empty
    const row = rows[0] || {};
    const existing = preferred.filter(c => c in row);
    return existing.length ? existing : preferred;
  }, [rows, selectedSet, idField]);

  const renderCell = (val: any) => {
    if (val == null) return '—';
    if (typeof val === 'string') return val.length > 120 ? val.slice(0, 117) + '…' : val;
    if (typeof val === 'number' || typeof val === 'boolean') return String(val);
    return '[object]';
  };

  const doCreate = async () => {
    setOpResult('');
    if (!(process.env.REACT_APP_ENABLE_DATAVERSE === 'true')) { setOpResult('Dataverse disabled'); return; }
    const body = parseJson(createBody);
    if (!body) { setOpResult('Invalid JSON'); return; }
    try {
      const t = await getDataverseToken();
      const res = await createRecord(selectedSet, t, body);
      if (res.ok) { setOpResult(`Created. id=${res.id || 'n/a'}`); await load(); }
      else setOpResult(`Create failed: ${res.status} ${res.text || ''}`);
    } catch (e: any) { setOpResult(`Create error: ${e?.message || e}`); }
  };

  const doUpdate = async () => {
    setOpResult('');
    const body = parseJson(updateBody);
    if (!body) { setOpResult('Invalid JSON'); return; }
    if (!guidLike(updateId)) { setOpResult('Provide a record GUID'); return; }
    try {
      const t = await getDataverseToken();
      const res = await updateRecord(selectedSet, updateId, t, body);
      if (res.ok) { setOpResult('Update OK'); await load(); }
      else setOpResult(`Update failed: ${res.status} ${res.text || ''}`);
    } catch (e: any) { setOpResult(`Update error: ${e?.message || e}`); }
  };

  const doDelete = async () => {
    setOpResult('');
    if (!guidLike(deleteId)) { setOpResult('Provide a record GUID'); return; }
    try {
      const t = await getDataverseToken();
      const res = await deleteRecord(selectedSet, deleteId, t);
      if (res.ok) { setOpResult('Delete OK'); await load(); }
      else setOpResult(`Delete failed: ${res.status} ${res.text || ''}`);
    } catch (e: any) { setOpResult(`Delete error: ${e?.message || e}`); }
  };

  return (
    <div>
      <h2 style={{ fontSize: 18, marginBottom: 12 }}>Dataverse Explorer</h2>
      <div className="small muted" style={{ marginBottom: 8 }}>Live only. All operations are against your Dataverse with your privileges. Use at your own risk.</div>

      <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12, flexWrap: 'wrap' }}>
        <div>
          <label className="small">Entity Set</label>
          <select value={selectedSet} onChange={e => setSelectedSet(e.target.value)} style={{ minWidth: 280, marginLeft: 8 }}>
            {sets.map(s => <SetOption key={s.set} name={s.name} value={s.set} />)}
          </select>
        </div>
        <div>
          <label className="small">Top</label>
          <input type="number" min={1} max={500} value={top} onChange={e => setTop(Math.max(1, Math.min(500, parseInt(e.target.value || '50'))))} style={{ width: 100, marginLeft: 8 }} />
        </div>
        <button className="btn sm" onClick={() => load()} disabled={loading}>{loading ? 'Loading…' : 'Refresh'}</button>
        <button className="btn ghost sm" onClick={() => { try { navigator.clipboard.writeText(pretty(rows)); window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Copied result JSON' } })); } catch {} }}>Copy JSON</button>
      </div>

      {error && <div className="small" style={{ color: '#b71c1c', marginBottom: 8 }}>Error: {error}</div>}

      {/* Grid */}
      <div style={{ border: '1px solid #eee', borderRadius: 6, overflowX: 'auto', marginBottom: 16 }}>
        <table className="small" style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              {columns.map(c => (
                <th key={c} style={{ textAlign: 'left', borderBottom: '1px solid #eee', padding: '8px 6px', whiteSpace: 'nowrap' }}>{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => (
              <tr key={i}>
                {columns.map(c => (
                  <td key={c} style={{ padding: '6px 6px', borderBottom: '1px solid #f6f6f6', maxWidth: 380 }}>{renderCell(r[c])}</td>
                ))}
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={columns.length} style={{ padding: 10 }} className="small muted">No rows</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {/* CRUD */}
      <div className="card" style={{ padding: 12, marginBottom: 12 }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>Create</div>
        <div className="small muted" style={{ marginBottom: 6 }}>
          Provide a JSON body. For lookups, use propertyName@odata.bind: <code>/{'{' }targetSet{ '}' }({'{' }GUID{ '}' })</code>.
        </div>
        <textarea value={createBody} onChange={e => setCreateBody(e.target.value)} style={{ width: '100%', minHeight: 120, fontFamily: 'monospace', fontSize: 12 }} />
        <div style={{ marginTop: 8 }}>
          <button className="btn sm" onClick={doCreate}>Create Record</button>
        </div>
      </div>

      <div className="card" style={{ padding: 12, marginBottom: 12 }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>Update</div>
        <div style={{ display: 'grid', gap: 8, gridTemplateColumns: '1fr' }}>
          <input placeholder="Record GUID" value={updateId} onChange={e => setUpdateId(e.target.value)} />
          <textarea value={updateBody} onChange={e => setUpdateBody(e.target.value)} style={{ width: '100%', minHeight: 120, fontFamily: 'monospace', fontSize: 12 }} />
          <div>
            <button className="btn sm" onClick={doUpdate}>Update Record</button>
          </div>
        </div>
      </div>

      <div className="card" style={{ padding: 12 }}>
        <div style={{ fontWeight: 700, marginBottom: 8 }}>Delete</div>
        <div style={{ display: 'flex', gap: 8 }}>
          <input placeholder="Record GUID" value={deleteId} onChange={e => setDeleteId(e.target.value)} style={{ flex: 1 }} />
          <button className="btn sm" onClick={doDelete}>Delete</button>
        </div>
      </div>

      {opResult && (
        <div className="small" style={{ marginTop: 10, whiteSpace: 'pre-wrap' }}>{opResult}</div>
      )}
    </div>
  );
};

export default DataverseExplorer;
