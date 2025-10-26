import React, { useEffect, useMemo, useState } from 'react';
import { useAuth as useAuthCtx } from '../../context/AuthContext';
import { showToast } from '../../utils/alerts';
import { getGraphToken } from '../../services/authTokens';
import { getApiBase } from '../../utils/runtimeConfig';
import {
  SharePointSite,
  SharePointDocumentLibrary,
  SharePointDocument,
  getSharePointSites,
  getDocumentLibraries,
  getDocuments,
  uploadFileToDrive,
  getFolderItems
} from '../../services/sharepointService';

// SharePoint Document Browser Component (extracted)
const SharePointBrowser: React.FC<{ onDocumentSelect: (docs: SharePointDocument[]) => void; canUpload?: boolean }> = ({ onDocumentSelect, canUpload = false }) => {
  const { getToken, login, account } = useAuthCtx();
  const [loading, setLoading] = useState(false);
  const [sites, setSites] = useState<SharePointSite[]>([]);
  const [selectedSite, setSelectedSite] = useState<string>('');
  const [libraries, setLibraries] = useState<SharePointDocumentLibrary[]>([]);
  const [selectedLibrary, setSelectedLibrary] = useState<string>('');
  const [documents, setDocuments] = useState<SharePointDocument[]>([]);
  const [selectedDocs, setSelectedDocs] = useState<Set<string>>(new Set());
  const [searchQuery, setSearchQuery] = useState('');
  const [spTab, setSpTab] = useState<'browse' | 'upload'>('browse');
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<number | null>(null);
  const [uploadStatuses, setUploadStatuses] = useState<{ name: string; progress: number; error?: string }[]>([]);
  const [folderItems, setFolderItems] = useState<any[]>([]);
  const [selectedFolderId, setSelectedFolderId] = useState<string>('root');
  const [breadcrumbs, setBreadcrumbs] = useState<Array<{ id: string; name: string }>>([{ id: 'root', name: 'Root' }]);
  const [spError, setSpError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [favoritesOnly, setFavoritesOnly] = useState(false);
  const [typeFilter, setTypeFilter] = useState<string>('all');
  const MAX_FILE_SIZE = 50 * 1024 * 1024;

  // Favorites handling
  const favKey = 'sp:favorites';
  const [favorites, setFavorites] = useState<Set<string>>(() => {
    try { const raw = localStorage.getItem(favKey); if (!raw) return new Set(); return new Set(JSON.parse(raw)); } catch { return new Set(); }
  });
  const persistFavs = (next: Set<string>) => { setFavorites(new Set(next)); try { localStorage.setItem(favKey, JSON.stringify(Array.from(next))); } catch {} };
  const docKey = (d: SharePointDocument) => d.id || d.webUrl || (d as any).name || '';
  const toggleFav = (d: SharePointDocument) => {
    const k = docKey(d);
    const next = new Set(favorites);
    const isAdding = !next.has(k);
    if (isAdding) next.add(k); else next.delete(k);
    persistFavs(next);
    try {
      if (d.id) {
        setSelectedDocs(prev => {
          const copy = new Set(prev);
          if (isAdding) copy.add(d.id as string); else copy.delete(d.id as string);
          return copy;
        });
      }
    } catch {}
  };
  const fileIcon = (name?: string) => {
    const n = (name || '').toLowerCase();
    if (n.endsWith('.pdf')) return '📕';
    if (n.endsWith('.doc') || n.endsWith('.docx')) return '📝';
    if (n.endsWith('.xls') || n.endsWith('.xlsx')) return '📊';
    if (n.endsWith('.ppt') || n.endsWith('.pptx')) return '📑';
    if (n.endsWith('.txt')) return '📄';
    if (n.endsWith('.html') || n.endsWith('.htm')) return '🌐';
    return '📁';
  };

  const loadSites = async () => {
    setLoading(true); setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const sitesData = await getSharePointSites(token);
      setSites(sitesData);
    } catch (error) {
      console.error('Failed to load SharePoint sites:', error);
      setSpError('Failed to load SharePoint sites. Ensure you are signed in and have granted Sites.Read.All and Files.Read.All.');
    } finally { setLoading(false); }
  };

  const loadLibraries = async (siteId: string) => {
    setLoading(true); setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const librariesData = await getDocumentLibraries(token, siteId);
      setLibraries(librariesData);
    } catch (error: any) {
      console.error('Failed to load document libraries:', error);
      const msg = typeof error?.message === 'string' ? error.message : '';
      setSpError(`Failed to load document libraries.${msg ? ' ' + msg : ''}`);
    } finally { setLoading(false); }
  };

  const loadDocuments = async (driveId: string, folderId: string = 'root') => {
    setLoading(true); setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const documentsData = await getDocuments(token, driveId, folderId, searchQuery);
      setDocuments(documentsData);
    } catch (error) {
      console.error('Failed to load documents:', error);
      setSpError('Failed to load documents.');
    } finally { setLoading(false); }
  };

  useEffect(() => { loadSites(); }, []);
  useEffect(() => { if (selectedSite) { loadLibraries(selectedSite); setSelectedLibrary(''); setDocuments([]); } }, [selectedSite]);
  useEffect(() => {
    if (selectedLibrary) {
      loadDocuments(selectedLibrary, 'root');
      (async () => {
        try {
          const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
          if (!token) throw new Error('No token available');
          const items = await getFolderItems(token, selectedLibrary, 'root');
          setFolderItems(items); setSelectedFolderId('root'); setBreadcrumbs([{ id: 'root', name: 'Root' }]);
        } catch (e) { console.error('Failed to load folder items', e); }
      })();
    }
  }, [selectedLibrary, searchQuery]);
  useEffect(() => { if (!selectedLibrary) return; if (spTab !== 'browse') return; loadDocuments(selectedLibrary, selectedFolderId || 'root'); }, [selectedFolderId, spTab]);
  useEffect(() => { const selected = Array.from(selectedDocs).map(id => documents.find(d => d.id === id)!).filter(Boolean); onDocumentSelect(selected); }, [selectedDocs, documents]);
  useEffect(() => { if (!canUpload && spTab === 'upload') setSpTab('browse'); }, [canUpload, spTab]);

  const toggleDocument = (docId: string) => {
    const next = new Set(selectedDocs);
    if (next.has(docId)) next.delete(docId); else next.add(docId);
    setSelectedDocs(next);
  };

  const navigateFolder = async (folderId: string, folderName: string) => {
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const items = await getFolderItems(token, selectedLibrary, folderId);
      setFolderItems(items); setSelectedFolderId(folderId);
      setBreadcrumbs(prev => { const idx = prev.findIndex(b => b.id === folderId); if (idx >= 0) return prev.slice(0, idx + 1); return [...prev, { id: folderId, name: folderName }]; });
    } catch (e) { console.error('Failed to navigate folder', e); }
  };

  const handleUpload = async (files: FileList | File[] | null) => {
    const arr: File[] = files ? Array.from(files as any) : [];
    if (!arr.length || !selectedLibrary) return;
    setUploading(true); setUploadProgress(0); setUploadStatuses([]);
    try {
      const token = await getToken(['Files.ReadWrite.All', 'Sites.Read.All']);
      if (!token) throw new Error('No token available');
      const uploadedDocs: SharePointDocument[] = [];
      for (let i = 0; i < arr.length; i++) {
        const f = arr[i]!;
        if ((f as any).size > MAX_FILE_SIZE) { setUploadStatuses(prev => [...prev, { name: f.name, progress: 0, error: `File exceeds ${Math.round(MAX_FILE_SIZE/1024/1024)}MB limit` }]); continue; }
        setUploadStatuses(prev => [...prev, { name: f.name, progress: 0 }]);
        try {
          const doc = await uploadFileToDrive(token, selectedLibrary, f, f.name, undefined, (p) => {
            setUploadProgress(p);
            setUploadStatuses(prev => { const copy = [...prev]; const idx = copy.findIndex(u => u.name === f.name); if (idx >= 0) copy[idx] = { ...copy[idx], progress: p }; return copy; });
          }, selectedFolderId);
          uploadedDocs.push(doc);
          setDocuments(prev => [{ id: doc.id, name: doc.name, webUrl: doc.webUrl, size: (doc as any).size || f.size, createdDateTime: (doc as any).createdDateTime || new Date().toISOString(), lastModifiedDateTime: (doc as any).lastModifiedDateTime || new Date().toISOString(), file: (doc as any).file || { mimeType: (f as any).type || 'application/octet-stream' }, parentReference: (doc as any).parentReference } as SharePointDocument, ...prev]);
          setSelectedDocs(prev => new Set(prev).add(doc.id));
        } catch (err: any) {
          const msg = typeof err?.message === 'string' ? err.message : 'Upload failed';
          setUploadStatuses(prev => { const copy = [...prev]; const idx = copy.findIndex(u => u.name === f.name); if (idx >= 0) copy[idx] = { ...copy[idx], error: msg }; return copy; });
        }
      }
      showToast(`Uploaded ${uploadedDocs.length} file(s)`, 'success');
      setSpTab('browse');
    } catch (e) { console.error('Upload failed', e); showToast('Upload failed', 'error'); }
    finally { setUploading(false); setUploadProgress(null); }
  };

  const onDropFiles = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault(); e.stopPropagation(); setIsDragging(false);
    const files = e.dataTransfer.files; if (files && files.length > 0) { void handleUpload(files); }
  };

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 16px 0', fontSize: 16 }}>SharePoint Documents</h3>
      <div style={{ marginBottom: 12 }}>
        {!account && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#fff3cd', padding: 8, borderRadius: 6, border: '1px solid #ffeeba' }}>
            <span>You're not signed in.</span>
            <button className="btn sm" onClick={() => login().then(() => loadSites())}>Sign in</button>
          </div>
        )}
        {spError && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#f8d7da', padding: 8, borderRadius: 6, border: '1px solid #f5c6cb', marginTop: 8 }}>
            <span style={{ flex: 1 }}>{spError}</span>
            <button className="btn ghost sm" onClick={() => loadSites()}>Retry</button>
          </div>
        )}
      </div>

      {/* Tabs */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, borderBottom: '1px solid #e0e0e0' }}>
        <button className={spTab === 'browse' ? 'btn sm' : 'btn ghost sm'} onClick={() => setSpTab('browse')}>Browse</button>
        {canUpload && (
          <button className={spTab === 'upload' ? 'btn sm' : 'btn ghost sm'} onClick={() => setSpTab('upload')}>Upload</button>
        )}
      </div>

      {loading && <div className="small muted">Loading...</div>}

      {/* Site Selection */}
      <div style={{ marginBottom: 16 }}>
        <label className="small">SharePoint Site:</label>
        <select value={selectedSite} onChange={e => setSelectedSite(e.target.value)} style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}>
          <option value="">Select a site...</option>
          {sites.map(site => <option key={site.id} value={site.id}>{site.displayName}</option>)}
        </select>
      </div>

      {/* Library Selection */}
      {selectedSite && (
        <div style={{ marginBottom: 16 }}>
          <label className="small">Document Library:</label>
          <select value={selectedLibrary} onChange={e => setSelectedLibrary(e.target.value)} style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}>
            <option value="">Select a library...</option>
            {libraries.map(lib => <option key={lib.id} value={lib.id}>{lib.displayName}</option>)}
          </select>
        </div>
      )}

      {/* Search (Browse Mode) */}
      {selectedLibrary && spTab === 'browse' && (
        <div style={{ marginBottom: 16 }}>
          <input type="text" placeholder="Search documents..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4 }} />
        </div>
      )}

      {/* Document List (Browse Mode) */}
      {spTab === 'browse' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          {selectedLibrary && (
            <div className="small" style={{ marginBottom: 8 }}>
              {breadcrumbs.map((b, i) => (
                <span key={b.id}>
                  {i > 0 && ' / '}
                  <a href="#" onClick={(e) => { e.preventDefault(); navigateFolder(b.id, b.name); }}>{b.name}</a>
                </span>
              ))}
            </div>
          )}
          {/* Filters & View options */}
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 8, flexWrap: 'wrap' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              <input type="checkbox" checked={favoritesOnly} onChange={e => setFavoritesOnly(e.target.checked)} />
              Favorites only
            </label>
            <select value={typeFilter} onChange={e => setTypeFilter(e.target.value)}>
              <option value="all">All types</option>
              <option value="pdf">PDF</option>
              <option value="word">Word</option>
              <option value="excel">Excel</option>
              <option value="ppt">PowerPoint</option>
              <option value="text">Text/HTML</option>
            </select>
          </div>
          {selectedLibrary && folderItems.filter(i => i.folder).map(f => (
            <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', borderBottom: '1px solid #f5f5f5' }}>
              <button className="btn ghost sm" onClick={() => navigateFolder(f.id, f.name)}>📁 {f.name}</button>
              <span className="small muted">{f.folder?.childCount ?? 0} items</span>
            </div>
          ))}
          {documents.length > 0 && (
            <>
              <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
                <button className="btn ghost sm" onClick={() => setSelectedDocs(new Set(documents.map(d => d.id)))}>Select All</button>
                <button className="btn ghost sm" onClick={() => setSelectedDocs(new Set())}>Clear</button>
                <span className="small muted">Selected: {selectedDocs.size}</span>
              </div>
              {documents
                .filter(d => !favoritesOnly || favorites.has(docKey(d)))
                .filter(d => {
                  const n = (d.name || '').toLowerCase();
                  if (typeFilter === 'all') return true;
                  if (typeFilter === 'pdf') return n.endsWith('.pdf');
                  if (typeFilter === 'word') return n.endsWith('.doc') || n.endsWith('.docx');
                  if (typeFilter === 'excel') return n.endsWith('.xls') || n.endsWith('.xlsx');
                  if (typeFilter === 'ppt') return n.endsWith('.ppt') || n.endsWith('.pptx');
                  if (typeFilter === 'text') return n.endsWith('.txt') || n.endsWith('.html') || n.endsWith('.htm');
                  return true;
                })
                .map(doc => (
                <div
                  key={doc.id}
                  onClick={() => toggleDocument(doc.id)}
                  role="button"
                  style={{ display: 'grid', gridTemplateColumns: 'auto auto 1fr auto', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0', cursor: 'pointer' }}
                >
                  <button className="btn ghost sm" title={favorites.has(docKey(doc)) ? 'Unpin' : 'Pin'} onClick={(e) => { e.stopPropagation(); toggleFav(doc); }}>{favorites.has(docKey(doc)) ? '⭐' : '☆'}</button>
                  <span aria-hidden>{fileIcon(doc.name)}</span>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{doc.name}</div>
                    <div className="small muted">{doc.size ? (doc.size / 1024).toFixed(1) + ' KB' : ''}{doc.lastModifiedDateTime ? ` • Modified ${new Date(doc.lastModifiedDateTime).toLocaleDateString()}` : ''}</div>
                    <a href={doc.webUrl} target="_blank" rel="noopener noreferrer" className="small" style={{ color: '#0066cc' }} onClick={(e) => e.stopPropagation()}>View in SharePoint ↗</a>
                  </div>
                  <input type="checkbox" checked={selectedDocs.has(doc.id)} onClick={(e) => e.stopPropagation()} onChange={() => toggleDocument(doc.id)} />
                </div>
              ))}
            </>
          )}
        </div>
      )}

      {/* Upload Mode */}
      {canUpload && spTab === 'upload' && (
        <div>
          {!selectedLibrary && <div className="small muted" style={{ marginBottom: 8 }}>Select a site and library to enable uploads.</div>}
          {selectedLibrary && (
            <div style={{ marginBottom: 12 }}>
              <label className="small">Target folder:</label>
              <select value={selectedFolderId} onChange={e => setSelectedFolderId(e.target.value)} style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}>
                <option value="root">/ (root)</option>
                {folderItems.filter(i => i.folder).map(f => (<option key={f.id} value={f.id}>/ {f.name}</option>))}
              </select>
            </div>
          )}
          <div onDragEnter={e => { e.preventDefault(); e.stopPropagation(); if (selectedLibrary && !uploading) setIsDragging(true); }} onDragOver={e => { e.preventDefault(); e.stopPropagation(); }} onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); }} onDrop={onDropFiles}
            style={{ border: '2px dashed ' + (isDragging ? 'var(--primary)' : '#ccc'), background: isDragging ? 'rgba(0,0,0,0.02)' : 'transparent', padding: 16, borderRadius: 8, textAlign: 'center', opacity: (!selectedLibrary || uploading) ? 0.6 : 1, pointerEvents: (!selectedLibrary || uploading) ? 'none' : 'auto' }}>
            <div className="small" style={{ marginBottom: 8 }}>Drag and drop files here</div>
            <div className="small muted">or</div>
            <div style={{ marginTop: 8 }}>
              <label className="btn ghost sm" style={{ cursor: (!selectedLibrary || uploading) ? 'not-allowed' : 'pointer' }}>
                Browse files
                <input type="file" multiple disabled={!selectedLibrary || uploading} accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.html" onChange={e => handleUpload(e.target.files)} style={{ display: 'none' }} />
              </label>
            </div>
            <div className="small muted" style={{ marginTop: 8 }}>Allowed: PDF, Word, Excel, PowerPoint, Text, HTML</div>
          </div>
          {uploading && (
            <div className="small" style={{ marginTop: 8 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Uploading...</span>
                <span>{uploadProgress ?? 0}%</span>
              </div>
              <div className="progressBar" aria-hidden="true" style={{ marginTop: 6 }}>
                <i style={{ width: `${uploadProgress ?? 0}%` }} />
              </div>
            </div>
          )}
          {uploadStatuses.length > 0 && (
            <div style={{ marginTop: 8 }}>
              {uploadStatuses.map((u, idx) => (
                <div key={idx} className="small" style={{ display: 'grid', gap: 4, marginBottom: 6 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
                    <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.name}</span>
                    {u.error ? (<span style={{ color: '#d33' }}>{u.error}</span>) : (<span>{u.progress}%</span>)}
                  </div>
                  {!u.error && (<div className="progressBar" aria-hidden="true"><i style={{ width: `${u.progress}%` }} /></div>)}
                </div>
              ))}
            </div>
          )}
          <div className="small muted" style={{ marginTop: 8 }}>Files are uploaded into the selected folder. Large files use a chunked upload session.</div>
        </div>
      )}
    </div>
  );
};

export default SharePointBrowser;
