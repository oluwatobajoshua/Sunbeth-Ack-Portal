import React, { useEffect, useState } from 'react';

export type SimpleDoc = {
  title: string;
  url: string;
  version?: number;
  requiresSignature?: boolean;
  driveId?: string;
  itemId?: string;
  source?: 'sharepoint' | 'url' | 'local';
  localFileId?: number | null;
  localUrl?: string | null;
};

export const DocumentListEditor: React.FC<{ onChange: (docs: SimpleDoc[]) => void; initial?: SimpleDoc[] }>
  = ({ onChange, initial = [] }) => {
  const [docs, setDocs] = useState<SimpleDoc[]>(initial);
  const [title, setTitle] = useState('');
  const [url, setUrl] = useState('');

  useEffect(() => { onChange(docs); }, [docs]);

  const addDoc = () => {
    const t = title.trim();
    const u = url.trim();
    if (!t || !u) return;
    setDocs(prev => [{ title: t, url: u, version: 1, requiresSignature: false }, ...prev]);
    setTitle(''); setUrl('');
  };
  const removeDoc = (idx: number) => setDocs(prev => prev.filter((_, i) => i !== idx));

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 16px 0', fontSize: 16 }}>Documents</h3>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 2fr auto', gap: 8, marginBottom: 12 }}>
        <input placeholder="Title" value={title} onChange={e => setTitle(e.target.value)} style={{ padding: 8, border: '1px solid #ddd', borderRadius: 4 }} />
        <input placeholder="URL (https://...)" value={url} onChange={e => setUrl(e.target.value)} style={{ padding: 8, border: '1px solid #ddd', borderRadius: 4 }} />
        <button className="btn sm" onClick={addDoc}>Add</button>
      </div>
      {docs.length === 0 && <div className="small muted">No documents added yet.</div>}
      {docs.length > 0 && (
        <div style={{ display: 'grid', gap: 8, maxHeight: 300, overflowY: 'auto' }}>
          {docs.map((d, idx) => (
            <div key={idx} style={{ display: 'grid', gridTemplateColumns: '1fr 3fr auto', gap: 8, alignItems: 'center' }}>
              <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.title}</div>
              <a href={d.url} target="_blank" rel="noopener noreferrer" className="small" style={{ color: '#0066cc', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.url}</a>
              <button className="btn ghost sm" onClick={() => removeDoc(idx)}>Remove</button>
            </div>
          ))}
        </div>
      )}
      <div className="small muted" style={{ marginTop: 8 }}>Tip: you can host files anywhere reachable (SharePoint, public storage, etc.). We store only metadata in SQLite.</div>
    </div>
  );
};
