/* eslint-disable max-lines-per-function */
import React, { useCallback, useEffect, useState } from 'react';
import { getApiBase } from '../../utils/runtimeConfig';
import type { SimpleDoc } from './DocumentListEditor';

const LocalLibraryPicker: React.FC<{ onAdd: (docs: SimpleDoc[]) => void }> = ({ onAdd }) => {
  const apiBase = (getApiBase() as string) || '';
  const [loading, setLoading] = useState(false);
  const [files, setFiles] = useState<Array<{ id: number; name: string; url: string; size?: number; uploadedAt?: string; mime?: string }>>([]);
  const [q, setQ] = useState('');
  const [selected, setSelected] = useState<Set<number>>(new Set());

  const load = useCallback(async () => {
    if (!apiBase) return;
    setLoading(true);
    try {
      const res = await fetch(`${apiBase}/api/library/list${q ? `?q=${encodeURIComponent(q)}` : ''}`);
      const j = await res.json();
      const arr = Array.isArray(j?.files) ? j.files : [];
      setFiles(arr.map((r: any) => ({ id: Number(r.id), name: String(r.name || 'file'), url: `${apiBase}${r.url}`, size: Number(r.size) || undefined, uploadedAt: r.uploadedAt || r.uploaded_at || undefined })));
    } catch {
      setFiles([]);
    } finally {
      setLoading(false);
    }
  }, [apiBase, q]);
  useEffect(() => { load(); }, [load]);

  const toggle = (id: number) => {
    setSelected(prev => { const n = new Set(prev); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  };
  const addSelected = () => {
    const chosen = files.filter(f => selected.has(f.id));
    if (chosen.length === 0) return;
    const docs: SimpleDoc[] = chosen.map(f => ({ title: f.name, url: f.url, version: 1, requiresSignature: false, source: 'local', localFileId: f.id, localUrl: f.url }));
    onAdd(docs);
    setSelected(new Set());
  };

  const fmtSize = (n?: number) => {
    if (!n || n <= 0) return '';
    if (n < 1024) return `${n} B`;
    if (n < 1024*1024) return `${(n/1024).toFixed(1)} KB`;
    return `${(n/1024/1024).toFixed(1)} MB`;
  };

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 12px 0', fontSize: 16 }}>Library (Server)</h3>
      <div className="small muted" style={{ marginBottom: 8 }}>Pick from previously saved files (recent first). Served from the app server for reliability.</div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr auto', gap: 8, marginBottom: 8 }}>
        <input placeholder="Search library..." value={q} onChange={e => setQ(e.target.value)} />
        <button className="btn ghost sm" onClick={load} disabled={loading}>Refresh</button>
      </div>
      {loading ? <div className="small muted">Loading...</div> : (
        files.length === 0 ? <div className="small muted">No files found.</div> : (
          <div style={{ maxHeight: 240, overflowY: 'auto', display: 'grid', gap: 6 }}>
            {files.map(f => (
              <label key={f.id} className="small" style={{ display: 'grid', gridTemplateColumns: 'auto 1fr auto', gap: 8, alignItems: 'center', padding: '6px 8px', borderBottom: '1px solid #f5f5f5' }}>
                <input type="checkbox" checked={selected.has(f.id)} onChange={() => toggle(f.id)} />
                <div style={{ overflow: 'hidden' }}>
                  <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{f.name}</div>
                  <div className="muted" style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                    <span>{fmtSize(f.size)}</span>
                    {f.uploadedAt && <span>• {new Date(f.uploadedAt).toLocaleString()}</span>}
                    <a href={f.url} target="_blank" rel="noreferrer">Preview ↗</a>
                  </div>
                </div>
                <span className="badge">local</span>
              </label>
            ))}
          </div>
        )
      )}
      <div style={{ marginTop: 8, textAlign: 'right' }}>
        <button className="btn sm" onClick={addSelected} disabled={selected.size === 0}>Add selected</button>
      </div>
    </div>
  );
};

export default LocalLibraryPicker;
