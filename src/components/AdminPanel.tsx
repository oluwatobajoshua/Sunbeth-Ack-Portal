import React, { useEffect, useState } from 'react';
import { useRBAC } from '../context/RBACContext';
// import { useAuth } from '../context/AuthContext';
import { GraphUser, GraphGroup, getUsers, getGroups, getOrganizationStructure, UserSearchFilters, getGroupMembers } from '../services/graphUserService';
import AnalyticsDashboard from './AnalyticsDashboard';
import Modal from './Modal';
import { sendEmail, buildBatchEmail /*, sendTeamsDirectMessage*/ } from '../services/notificationService';
import { getGraphToken } from '../services/authTokens';
import { runAuthAndGraphCheck, Step } from '../diagnostics/health';
import { getBusinesses, createBusiness, updateBusiness, deleteBusiness } from '../services/dbService';
// SharePoint Lists removed; SQLite-only mode
// SharePoint document browsing & upload
import { SharePointSite, SharePointDocumentLibrary, SharePointDocument, getSharePointSites, getDocumentLibraries, getDocuments, uploadFileToDrive, getFolderItems } from '../services/sharepointService';
import BatchCreationDebug from './BatchCreationDebug';

// Enhanced Admin Settings Component
type AdminSettingsProps = { canEdit: boolean };

const AdminSettings: React.FC<AdminSettingsProps> = ({ canEdit }) => {
  const { account } = useAuthCtx();
  const storageKey = 'admin_settings';
  const [settings, setSettings] = useState({
    enableUpload: false,
    requireSig: false,
    autoReminder: true,
    reminderDays: 3,
    allowBulkAssignment: true,
    requireApproval: false
  });


  useEffect(() => {
    try {
      const raw = localStorage.getItem(storageKey);
      if (raw) {
        const obj = JSON.parse(raw);
        setSettings({ ...settings, ...obj });
      }
    } catch {}
  }, []);

  const apply = () => {
    if (!canEdit) return;
    try {
      localStorage.setItem(storageKey, JSON.stringify(settings));
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Settings saved' } }));
    } catch (e) {
      console.warn(e);
    }
  };

  const seedSqliteForMe = async () => {
    try {
      if (!((process.env.REACT_APP_ENABLE_SQLITE === 'true') && process.env.REACT_APP_API_BASE)) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable SQLite (REACT_APP_ENABLE_SQLITE=true) and set REACT_APP_API_BASE to seed.' } }));
        return;
      }
      if (!account?.username) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Sign in first to seed data for your account.' } }));
        return;
      }
      const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
      const res = await fetch(`${base}/api/seed?email=${encodeURIComponent(account.username)}`, { method: 'POST' });
      if (!res.ok) throw new Error('Seed failed');
      const j = await res.json().catch(() => ({}));
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `SQLite seeded. BatchId: ${j?.batchId ?? 'n/a'}` } }));
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'SQLite seed failed' } }));
    }
  };



  const grantCorePermissions = async () => {
    try {
      // Request common scopes used across the Admin feature set
      // Do these in series to present clearer consent prompts
      await getGraphToken(['User.Read']);
      await getGraphToken(['Group.Read.All']);
      await getGraphToken(['Sites.Read.All','Files.ReadWrite.All']);
      await getGraphToken(['Mail.Send']);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Core Graph permissions granted (if consented)' } }));
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Grant permissions failed' } }));
    }
  };



  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <h3 style={{ margin: 0, fontSize: 16 }}>System Settings</h3>
      
      <div className="grid" style={{ gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.enableUpload} onChange={e => setSettings({...settings, enableUpload: e.target.checked})} disabled={!canEdit} />
          <span className="small">Enable document upload</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.requireSig} onChange={e => setSettings({...settings, requireSig: e.target.checked})} disabled={!canEdit} />
          <span className="small">Require digital signatures</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.autoReminder} onChange={e => setSettings({...settings, autoReminder: e.target.checked})} disabled={!canEdit} />
          <span className="small">Auto-send reminders</span>
        </label>
        
        <label style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <input type="checkbox" checked={settings.allowBulkAssignment} onChange={e => setSettings({...settings, allowBulkAssignment: e.target.checked})} disabled={!canEdit} />
          <span className="small">Allow bulk assignments</span>
        </label>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <span className="small">Reminder frequency:</span>
        <select value={settings.reminderDays} onChange={e => setSettings({...settings, reminderDays: parseInt(e.target.value)})} disabled={!canEdit}>
          <option value={1}>Daily</option>
          <option value={3}>Every 3 days</option>
          <option value={7}>Weekly</option>
          <option value={14}>Bi-weekly</option>
        </select>
      </div>

      <div style={{ display: 'flex', gap: 8, marginTop: 8, alignItems: 'center', flexWrap: 'wrap' }}>
        {canEdit && <button className="btn" onClick={apply}>Save Settings</button>}
        {!canEdit && <span className="small muted">Read-only access</span>}
        {canEdit && <button className="btn ghost" onClick={seedSqliteForMe} title="Seed SQLite with a demo batch, docs, and recipients for your account">Seed SQLite (for me)</button>}
        {canEdit && <button className="btn ghost" onClick={grantCorePermissions} title="Request common Microsoft Graph permissions in one go">Grant Core Permissions</button>}
      </div>

      {/* SharePoint provisioning UI removed */}
    </div>
  );
};

// User/Group Selection Component
const UserGroupSelector: React.FC<{ onSelectionChange: (selection: any) => void }> = ({ onSelectionChange }) => {
  const { getToken, login, account } = useAuthCtx();
  const [loading, setLoading] = useState(false);
  const [hadError, setHadError] = useState<string | null>(null);
  const [tab, setTab] = useState<'users' | 'groups' | 'structure'>('users');
  const [users, setUsers] = useState<GraphUser[]>([]);
  const [groups, setGroups] = useState<GraphGroup[]>([]);
  const [orgStructure, setOrgStructure] = useState<{ departments: string[]; jobTitles: string[]; locations: string[] }>({ departments: [], jobTitles: [], locations: [] });
  const [filters, setFilters] = useState<UserSearchFilters>({});
  const [localSearch, setLocalSearch] = useState<string>('');
  const [selectedUsers, setSelectedUsers] = useState<Set<string>>(new Set());
  const [selectedGroups, setSelectedGroups] = useState<Set<string>>(new Set());
  const [usersPage, setUsersPage] = useState<number>(1);
  const [groupsPage, setGroupsPage] = useState<number>(1);
  const pageSize = 50;

  const loadData = async () => {
    setLoading(true);
    setHadError(null);
    try {
      const token = await getToken(['User.Read.All', 'Group.Read.All']);
      if (!token) throw new Error('No token available');

      const [usersData, groupsData, structureData] = await Promise.all([
        getUsers(token, filters),
        getGroups(token),
        getOrganizationStructure(token)
      ]);

      setUsers(usersData);
      setGroups(groupsData);
      setOrgStructure(structureData);
    } catch (error: any) {
      console.error('Failed to load user/group data:', error);
      const msg = typeof error?.message === 'string' ? error.message : '';
      const hint = msg.includes('No active account')
        ? 'Please sign in to continue.'
        : 'Ask your admin to grant Microsoft Graph permissions (User.Read.All and Group.Read.All) to this app.';
      setHadError(`${msg || 'Failed to load user data.'} ${hint}`.trim());
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Failed to load user data. ${hint}` } }));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
    setUsersPage(1);
    setGroupsPage(1);
  }, [filters]);

  // Debounce search input before applying to filters
  useEffect(() => {
    const h = setTimeout(() => {
      setFilters(prev => ({ ...prev, search: localSearch || undefined }));
    }, 300);
    return () => clearTimeout(h);
  }, [localSearch]);

  useEffect(() => {
    onSelectionChange({
      users: Array.from(selectedUsers).map(id => users.find(u => u.id === id)!).filter(Boolean),
      groups: Array.from(selectedGroups).map(id => groups.find(g => g.id === id)!).filter(Boolean)
    });
  }, [selectedUsers, selectedGroups, users, groups]);

  const toggleUser = (userId: string) => {
    const newSelection = new Set(selectedUsers);
    if (newSelection.has(userId)) {
      newSelection.delete(userId);
    } else {
      newSelection.add(userId);
    }
    setSelectedUsers(newSelection);
  };

  const toggleGroup = (groupId: string) => {
    const newSelection = new Set(selectedGroups);
    if (newSelection.has(groupId)) {
      newSelection.delete(groupId);
    } else {
      newSelection.add(groupId);
    }
    setSelectedGroups(newSelection);
  };

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 16px 0', fontSize: 16 }}>Assign to Users & Groups</h3>
      <div style={{ marginBottom: 12 }}>
        {!account && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#fff3cd', padding: 8, borderRadius: 6, border: '1px solid #ffeeba' }}>
            <span>You're not signed in.</span>
            <button className="btn sm" onClick={() => login().then(() => loadData())}>Sign in</button>
          </div>
        )}
        {hadError && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#f8d7da', padding: 8, borderRadius: 6, border: '1px solid #f5c6cb', marginTop: 8 }}>
            <span style={{ flex: 1 }}>{hadError}</span>
            <button className="btn ghost sm" onClick={() => loadData()}>Retry</button>
          </div>
        )}
      </div>
      
      {/* Tab Navigation */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, borderBottom: '1px solid #e0e0e0' }}>
        {(['users', 'groups', 'structure'] as const).map(t => (
          <button 
            key={t}
            className={tab === t ? 'btn sm' : 'btn ghost sm'}
            onClick={() => setTab(t)}
          >
            {t === 'users' ? `Users (${users.length})` : t === 'groups' ? `Groups (${groups.length})` : 'Filters'}
          </button>
        ))}
      </div>

      {loading && <div className="small muted">Loading...</div>}

      {/* Filters Tab */}
      {tab === 'structure' && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <div>
            <label className="small">Search:</label>
            <input 
              type="text" 
              placeholder="Name, email..." 
              value={localSearch}
              onChange={e => setLocalSearch(e.target.value)}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            />
          </div>
          <div>
            <label className="small">Department:</label>
            <select 
              value={filters.department || ''} 
              onChange={e => setFilters({...filters, department: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Departments</option>
              {orgStructure.departments.map(dept => <option key={dept} value={dept}>{dept}</option>)}
            </select>
          </div>
          <div>
            <label className="small">Job Title:</label>
            <select 
              value={filters.jobTitle || ''} 
              onChange={e => setFilters({...filters, jobTitle: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Titles</option>
              {orgStructure.jobTitles.map(title => <option key={title} value={title}>{title}</option>)}
            </select>
          </div>
          <div>
            <label className="small">Location:</label>
            <select 
              value={filters.location || ''} 
              onChange={e => setFilters({...filters, location: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Locations</option>
              {orgStructure.locations.map(loc => <option key={loc} value={loc}>{loc}</option>)}
            </select>
          </div>
        </div>
      )}

      {/* Users Tab */}
      {tab === 'users' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
            <button className="btn ghost sm" onClick={() => setSelectedUsers(new Set(users.map(u => u.id)))}>Select All</button>
            <button className="btn ghost sm" onClick={() => setSelectedUsers(new Set())}>Clear</button>
            <span className="small muted">Selected: {selectedUsers.size}</span>
          </div>
          {users.slice(0, usersPage * pageSize).map(user => (
            <div key={user.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0' }}>
              <input 
                type="checkbox" 
                checked={selectedUsers.has(user.id)} 
                onChange={() => toggleUser(user.id)} 
              />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 500 }}>{user.displayName}</div>
                <div className="small muted">{user.userPrincipalName}</div>
                {user.department && <div className="small muted">{user.department} • {user.jobTitle}</div>}
              </div>
            </div>
          ))}
          {(usersPage * pageSize) < users.length && (
            <div style={{ padding: 8, textAlign: 'center' }}>
              <button className="btn ghost sm" onClick={() => setUsersPage(p => p + 1)}>Load more</button>
              <div className="small muted" style={{ marginTop: 6 }}>{Math.min(usersPage * pageSize, users.length)} of {users.length}</div>
            </div>
          )}
        </div>
      )}

      {/* Groups Tab */}
      {tab === 'groups' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
            <button className="btn ghost sm" onClick={() => setSelectedGroups(new Set(groups.map(g => g.id)))}>Select All</button>
            <button className="btn ghost sm" onClick={() => setSelectedGroups(new Set())}>Clear</button>
            <span className="small muted">Selected: {selectedGroups.size}</span>
          </div>
          {groups.slice(0, groupsPage * pageSize).map(group => (
            <div key={group.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0' }}>
              <input 
                type="checkbox" 
                checked={selectedGroups.has(group.id)} 
                onChange={() => toggleGroup(group.id)} 
              />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 500 }}>{group.displayName}</div>
                {group.description && <div className="small muted">{group.description}</div>}
                <div className="small muted">{group.memberCount || 0} members</div>
              </div>
            </div>
          ))}
          {(groupsPage * pageSize) < groups.length && (
            <div style={{ padding: 8, textAlign: 'center' }}>
              <button className="btn ghost sm" onClick={() => setGroupsPage(p => p + 1)}>Load more</button>
              <div className="small muted" style={{ marginTop: 6 }}>{Math.min(groupsPage * pageSize, groups.length)} of {groups.length}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// Simple Document List Editor (SQLite-only)
import { useAuth as useAuthCtx } from '../context/AuthContext';
type SimpleDoc = { title: string; url: string; version?: number; requiresSignature?: boolean; driveId?: string; itemId?: string; source?: 'sharepoint' | 'url' };
const DocumentListEditor: React.FC<{ onChange: (docs: SimpleDoc[]) => void; initial?: SimpleDoc[] }>
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

// SharePoint Document Browser Component (restored)
const SharePointBrowser: React.FC<{ onDocumentSelect: (docs: SharePointDocument[]) => void }> = ({ onDocumentSelect }) => {
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
  const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB UX guard

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
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Uploaded ${uploadedDocs.length} file(s)` } }));
      setSpTab('browse');
    } catch (e) { console.error('Upload failed', e); window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Upload failed' } })); }
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
        {(['browse', 'upload'] as const).map(t => (
          <button key={t} className={spTab === t ? 'btn sm' : 'btn ghost sm'} onClick={() => setSpTab(t)}>
            {t === 'browse' ? 'Browse' : 'Upload'}
          </button>
        ))}
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
              {documents.map(doc => (
                <div key={doc.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0' }}>
                  <input type="checkbox" checked={selectedDocs.has(doc.id)} onChange={() => toggleDocument(doc.id)} />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 500 }}>{doc.name}</div>
                    <div className="small muted">{doc.size ? (doc.size / 1024).toFixed(1) + ' KB' : ''}{doc.lastModifiedDateTime ? ` • Modified ${new Date(doc.lastModifiedDateTime).toLocaleDateString()}` : ''}</div>
                    <a href={doc.webUrl} target="_blank" rel="noopener noreferrer" className="small" style={{ color: '#0066cc' }}>View in SharePoint ↗</a>
                  </div>
                </div>
              ))}
            </>
          )}
        </div>
      )}

      {/* Upload Mode */}
      {spTab === 'upload' && (
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

// Main Admin Panel Component
const AdminPanel: React.FC = () => {
  const { role, canSeeAdmin, canEditAdmin } = useRBAC();
  const { account } = useAuthCtx();
  const [activeTab, setActiveTab] = useState<'overview' | 'settings' | 'manage' | 'batch' | 'analytics'>('overview');
  const [editingBatchId, setEditingBatchId] = useState<string | null>(null);
  const [originalRecipientEmails, setOriginalRecipientEmails] = useState<Set<string>>(new Set());
  const [originalDocUrls, setOriginalDocUrls] = useState<Set<string>>(new Set());
  const [apiHealth, setApiHealth] = useState<'unknown' | 'ok' | 'down'>('unknown');
  const pingApi = async () => {
    try {
      if (!sqliteEnabled) { setApiHealth('unknown'); return; }
      const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
      const res = await fetch(`${base}/api/health`, { cache: 'no-store' });
      setApiHealth(res.ok ? 'ok' : 'down');
    } catch { setApiHealth('down'); }
  };
  useEffect(() => { pingApi(); }, [/* on mount and when sqlite flag changes */]);
  const [healthOpen, setHealthOpen] = useState(false);
  const [healthSteps, setHealthSteps] = useState<Step[] | null>(null);
  const [granting, setGranting] = useState(false);
  const [permStatus, setPermStatus] = useState<Record<string, boolean>>({});

  const MODAL_TOGGLE_KEY = 'sunbeth:admin:useModalSelectors';
  const adminLight = (process.env.REACT_APP_ADMIN_LIGHT || '').toLowerCase() === 'true';
  const defaultModalToggle = ((): boolean => {
    const env = (process.env.REACT_APP_ADMIN_MODAL_SELECTORS || '').toLowerCase();
    if (env === 'true') return true; if (env === 'false') return false; return true; // default ON to avoid mounting heavy selectors
  })();
  const [useModalSelectors, setUseModalSelectors] = useState<boolean>(() => {
    try { const v = localStorage.getItem(MODAL_TOGGLE_KEY); if (v === 'true') return true; if (v === 'false') return false; } catch {}
    return defaultModalToggle;
  });
  const [showDebugConsole, setShowDebugConsole] = useState(false);
  const [usersModalOpen, setUsersModalOpen] = useState(false);
  const [docsModalOpen, setDocsModalOpen] = useState(false);

  const requiredScopes = ['User.Read','User.Read.All','Group.Read.All','Mail.Send'];

  const checkPermissions = async () => {
    const status: Record<string, boolean> = {};
    for (const scope of requiredScopes) {
      try { await getGraphToken([scope]); status[scope] = true; }
      catch { status[scope] = false; }
    }
    setPermStatus(status);
  };

  useEffect(() => { if (!adminLight) { checkPermissions().catch(() => {}); } }, [adminLight]);

  const [batchForm, setBatchForm] = useState<{
    name: string;
    startDate: string;
    dueDate: string;
    description: string;
    selectedUsers: GraphUser[];
    selectedGroups: GraphGroup[];
    selectedDocuments: SimpleDoc[];
    notifyByEmail: boolean;
    notifyByTeams: boolean;
  }>({
    name: '',
    startDate: '',
    dueDate: '',
    description: '',
    selectedUsers: [],
    selectedGroups: [],
    selectedDocuments: [],
    notifyByEmail: true,
    notifyByTeams: false
  });
  // Users to display in Business Mapping (includes selected users and optionally expanded group members)
  const [mappingUsers, setMappingUsers] = useState<GraphUser[]>([]);
  useEffect(() => {
    // Reset mapping list to explicitly selected users when selection changes
    setMappingUsers(batchForm.selectedUsers);
  }, [batchForm.selectedUsers]);
  const expandGroupsForMapping = async () => {
    try {
      if (batchForm.selectedGroups.length === 0) return;
      const token = await getGraphToken(['Group.Read.All','User.Read']);
      const membersArrays = await Promise.all(
        batchForm.selectedGroups.map(g => getGroupMembers(token, g.id).catch(() => []))
      );
      const members = ([] as GraphUser[]).concat(...membersArrays);
      // Merge selected users + group members (unique by email lower)
      const mergedByEmail = new Map<string, GraphUser>();
      const push = (u: GraphUser) => {
        const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
        if (!email) return;
        if (!mergedByEmail.has(email)) mergedByEmail.set(email, u);
      };
      batchForm.selectedUsers.forEach(push);
      members.forEach(push);
      setMappingUsers(Array.from(mergedByEmail.values()));
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Loaded ${members.length} member(s) from selected group(s)` } }));
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to expand groups for mapping' } }));
    }
  };
  // Maintain user -> business mapping by email
  const setUserBusiness = (emailOrUpn: string, businessId: number | null) => {
    const key = (emailOrUpn || '').trim().toLowerCase();
    if (!key) return;
    setBusinessMap(prev => ({ ...prev, [key]: businessId }));
  };
  const applyBusinessToAll = (businessId: number | null) => {
    const next: Record<string, number | null> = {};
    for (const u of mappingUsers) {
      const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
      if (email) next[email] = businessId;
    }
    setBusinessMap(next);
  };



  if (!canSeeAdmin) {
    return (
      <div className="container">
        <div className="card">
          <h2 style={{ color: '#d63384', margin: 0 }}>Access Denied</h2>
          <p>You do not have permission to access the Admin Panel.</p>
          <p className="small muted">Current role: {role}</p>
        </div>
      </div>
    );
  }



  const tabs = [
    { id: 'overview', label: 'Overview', icon: '📊' },
    { id: 'settings', label: 'Settings', icon: '⚙️' },
    { id: 'manage', label: 'Manage', icon: '🧰' },
    { id: 'batch', label: 'Create Batch', icon: '📝' },
    { id: 'analytics', label: 'Analytics', icon: '📈' }
  ];
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  const [overviewStats, setOverviewStats] = useState<{ totalBatches: number; activeBatches: number; totalUsers: number; completionRate: number; overdueBatches: number; avgCompletionTime: number } | null>(null);
  type Business = { id: number; name: string; code?: string; isActive?: boolean };
  const [businesses, setBusinesses] = useState<Business[]>([]);
  const [businessMap, setBusinessMap] = useState<Record<string, number | null>>({}); // emailLower -> businessId
  const [defaultBusinessId, setDefaultBusinessId] = useState<number | ''>('');
  useEffect(() => {
    if (!sqliteEnabled) return;
    (async () => {
      try {
        const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
        const res = await fetch(`${base}/api/stats`);
        if (!res.ok) throw new Error('stats_failed');
        const j = await res.json();
        setOverviewStats(j);
      } catch {}
    })();
  }, [sqliteEnabled]);
  // Load businesses from SQLite
  useEffect(() => {
    if (!sqliteEnabled) return;
    (async () => {
      try {
        const businessesData = await getBusinesses();
        const arr: any[] = Array.isArray(businessesData) ? businessesData : [];
        // Normalize shape
        const mapped: Business[] = arr
          .map((row: any) => ({
            id: Number(row.id ?? row.businessId ?? row.ID ?? row.toba_businessid),
            name: String(row.name ?? row.Title ?? row.title ?? row.toba_name ?? row.code ?? 'Business'),
            code: row.code ?? row.toba_code,
            isActive: (row.isActive ?? row.toba_isactive) ? true : false
          } as Business))
          .filter((b: Business) => Number.isFinite(b.id) && !!b.name);
        setBusinesses(mapped);
      } catch (e) {
        console.warn('Failed to load businesses', e);
        setBusinesses([]);
      }
    })();
  }, [sqliteEnabled]);

  const saveBatch = async () => {
    try {
      // SQLite-only persistence via API
      if (!((process.env.REACT_APP_ENABLE_SQLITE === 'true') && process.env.REACT_APP_API_BASE)) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable SQLite (REACT_APP_ENABLE_SQLITE=true) and set REACT_APP_API_BASE.' } }));
        return;
      }
      const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');

      // Build recipients from selected users and expand selected groups into members
      const recipientSet = new Map<string, { address: string; name?: string }>();
      // Track origins (user/group) to apply business defaults
      const recipientOrigins = new Map<string, Set<string>>(); // emailLower -> Set of groupIds
      const addRecipient = (addrRaw: string, name?: string, originGroupId?: string) => {
        const addr = (addrRaw || '').trim();
        if (!addr) return;
        const key = addr.toLowerCase();
        if (!recipientSet.has(key)) recipientSet.set(key, { address: addr, name });
        if (originGroupId) {
          const set = recipientOrigins.get(key) || new Set<string>();
          set.add(originGroupId);
          recipientOrigins.set(key, set);
        }
      };

      for (const u of batchForm.selectedUsers) {
        addRecipient((u.mail || u.userPrincipalName || ''), u.displayName);
      }

      if (batchForm.selectedGroups.length > 0) {
        try {
          const token = await getGraphToken(['Group.Read.All', 'User.Read']);
          const membersArrays = await Promise.all(
            batchForm.selectedGroups.map(g => getGroupMembers(token, g.id).then(ms => ({ gid: g.id, members: ms })).catch(() => ({ gid: g.id, members: [] })))
          );
          for (const { gid, members } of membersArrays) {
            for (const m of members) {
              addRecipient((m.mail || m.userPrincipalName || ''), m.displayName, gid);
            }
          }
        } catch (e) {
          console.warn('Failed to expand group members for notifications', e);
        }
      }

      const recipients = Array.from(recipientSet.values());

      // Helper maps for extra profile info
      const userByEmailLower = new Map<string, GraphUser>();
      for (const u of batchForm.selectedUsers) {
        const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
        if (email) userByEmailLower.set(email, u);
      }

      // Build email content
      const { subject, bodyHtml } = buildBatchEmail({
        appUrl: window.location.origin,
        batchName: batchForm.name,
        startDate: batchForm.startDate,
        dueDate: batchForm.dueDate,
        description: batchForm.description
      });

      // Determine which recipients should be notified (only new ones during edit)
      let recipientsToNotify = recipients;
      if (editingBatchId) {
        const isNew = (addr: string) => !originalRecipientEmails.has((addr || '').trim().toLowerCase());
        const filtered = recipients.filter(r => isNew(r.address));
        recipientsToNotify = filtered;
      }
      if (batchForm.notifyByEmail && recipientsToNotify.length > 0) {
        await sendEmail(recipientsToNotify, subject, bodyHtml);
      }
      // Teams optional (requires Chat.ReadWrite)
      // if (batchForm.notifyByTeams) {
      //   const userIds = batchForm.selectedUsers.map(u => u.id);
      //   await sendTeamsDirectMessage(userIds, `New acknowledgement assigned: ${batchForm.name}`);
      // }

      // 1) Create or update batch in SQLite
      let batchId: string | undefined;
      if (!editingBatchId) {
        const createRes = await fetch(`${base}/api/batches`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: batchForm.name,
            startDate: batchForm.startDate || null,
            dueDate: batchForm.dueDate || null,
            description: batchForm.description || null,
            status: 1
          })
        });
        if (!createRes.ok) throw new Error('batch_create_failed');
        const createJson = await createRes.json();
        const batchIdRaw = (createJson?.id ?? createJson?.batchId ?? createJson?.toba_batchid ?? createJson?.ID);
        batchId = typeof batchIdRaw === 'string' ? batchIdRaw : (Number.isFinite(Number(batchIdRaw)) ? String(batchIdRaw) : undefined);
        
        console.log('🔍 DEBUG: Batch creation result:', {
          createJson,
          batchIdRaw,
          finalBatchId: batchId
        });
        
        if (!batchId) throw new Error('batch_id_missing');
        
        // Small delay to ensure batch is fully committed before adding relations
        await new Promise(resolve => setTimeout(resolve, 100));
      } else {
        batchId = editingBatchId;
        const updateRes = await fetch(`${base}/api/batches/${encodeURIComponent(batchId)}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: batchForm.name,
            startDate: batchForm.startDate || null,
            dueDate: batchForm.dueDate || null,
            description: batchForm.description || null,
            status: 1
          })
        });
        if (!updateRes.ok) throw new Error('batch_update_failed');
      }

      // 2) Add documents (only new ones on edit; full set on create)
      const allDocsPayload = batchForm.selectedDocuments.map(d => ({
        title: d.title,
        url: d.url,
        version: d.version ?? 1,
        requiresSignature: !!d.requiresSignature,
        driveId: (d as any).driveId || null,
        itemId: (d as any).itemId || null,
        source: (d as any).source || null
      }));
      const docsToPost = !editingBatchId
        ? allDocsPayload
        : allDocsPayload.filter(d => !originalDocUrls.has((d.url || '').trim()));
      
      console.log('🔍 DEBUG: Documents to post:', {
        isCreating: !editingBatchId,
        totalDocs: allDocsPayload.length,
        docsToPost: docsToPost.length,
        batchId
      });
      
      if (docsToPost.length > 0) {
        const docsRes = await fetch(`${base}/api/batches/${batchId}/documents`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ documents: docsToPost })
        });
        if (!docsRes.ok) {
          const errorText = await docsRes.text().catch(() => '');
          console.error('Documents insert failed:', docsRes.status, errorText);
          throw new Error(`docs_insert_failed: ${docsRes.status} - ${errorText}`);
        } else {
          const docsResult = await docsRes.json().catch(() => null);
          console.log('✅ DEBUG: Documents API success:', docsResult);
        }
      }

      // 3) Add recipients
      const recipientsPayloadAll = recipients.map(r => {
        const emailLower = (r.address || '').toLowerCase();
        const u = userByEmailLower.get(emailLower);
        let primaryGroupName: string | undefined = undefined;
        const origins = recipientOrigins.get(emailLower);
        if (origins && origins.size > 0) {
          const firstGid = origins.values().next().value as string;
          const g = batchForm.selectedGroups.find(x => x.id === firstGid);
          if (g?.displayName) primaryGroupName = g.displayName;
        }
        const mappedBusinessId = (businessMap[emailLower] ?? (defaultBusinessId !== '' ? Number(defaultBusinessId) : null));
        return {
          businessId: mappedBusinessId,
          user: emailLower,
          email: emailLower,
          userEmail: emailLower,
          userPrincipalName: emailLower,
          displayName: r.name || undefined,
          department: u?.department || undefined,
          jobTitle: u?.jobTitle || undefined,
          location: u?.officeLocation || undefined,
          primaryGroup: primaryGroupName || undefined
        };
      });
      const recipientsPayload = editingBatchId
        ? recipientsPayloadAll.filter(r => !originalRecipientEmails.has((r.email || '').trim().toLowerCase()))
        : recipientsPayloadAll;
      
      console.log('🔍 DEBUG: Recipients to post:', {
        isCreating: !editingBatchId,
        totalRecipients: recipientsPayloadAll.length,
        recipientsToPost: recipientsPayload.length,
        batchId,
        sampleRecipient: recipientsPayload[0]
      });
      
      if (recipientsPayload.length > 0) {
        const recRes = await fetch(`${base}/api/batches/${batchId}/recipients`, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ recipients: recipientsPayload })
        });
        if (!recRes.ok) {
          const errorText = await recRes.text().catch(() => '');
          console.error('Recipients insert failed:', recRes.status, errorText);
          throw new Error(`recipients_insert_failed: ${recRes.status} - ${errorText}`);
        } else {
          const recResult = await recRes.json().catch(() => null);
          console.log('✅ DEBUG: Recipients API success:', recResult);
        }
        try {
          const verify = await fetch(`${base}/api/batches/${batchId}/recipients`, { cache: 'no-store' });
          const rows = await verify.json();
          if (!Array.isArray(rows) || rows.length === 0) {
            console.warn('Recipients verification returned empty for batch', batchId);
            window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Warning: recipients not linked to batch (verification empty)' } }));
          }
        } catch (e) {
          console.warn('Recipients verification failed', e);
        }
      }

      // Final feedback
  const actionWord = editingBatchId ? 'updated' : 'created';
  window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Batch "${batchForm.name}" ${actionWord} in SQLite${batchForm.notifyByEmail ? ' and email sent' : ''}.` } }));

      // Reset form
      setBatchForm({
        name: '',
        startDate: '',
        dueDate: '',
        description: '',
        selectedUsers: [],
        selectedGroups: [],
        selectedDocuments: [],
        notifyByEmail: true,
        notifyByTeams: false
      });
  setBusinessMap({});
  setDefaultBusinessId('');
  setEditingBatchId(null);
  setOriginalRecipientEmails(new Set());
  setOriginalDocUrls(new Set());

    } catch (e) {
      console.error('Save batch failed', e);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to save batch or send notifications' } }));
    }
  };

  // Load an existing batch into the form for editing
  const startEditBatch = async (id: string) => {
    try {
      if (!sqliteEnabled) return;
      const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
      // Fetch all batches and find one
      const bRes = await fetch(`${base}/api/batches`);
      const all = await bRes.json();
      const b = (Array.isArray(all) ? all : []).find((x: any) => String(x.toba_batchid || x.id || x.batchId || x.ID) === String(id));
      if (!b) throw new Error('batch_not_found');
      // Documents
  let docs: any[] = [];
      try {
        const dRes = await fetch(`${base}/api/batches/${encodeURIComponent(id)}/documents`);
        docs = await dRes.json();
      } catch { docs = []; }
      // Recipients
  let recs: any[] = [];
      try {
        const rRes = await fetch(`${base}/api/recipients`);
        const allRecs = await rRes.json();
        recs = (Array.isArray(allRecs) ? allRecs : []).filter((r: any) => String(r.batchId) === String(id));
      } catch { recs = []; }

      // Build form
      const selectedDocuments = (docs || []).map((d: any) => ({
        title: d.title || d.name || d.toba_title || 'Document',
        url: d.url || d.webUrl || d.toba_fileurl,
        version: Number(d.version || d.toba_version || 1),
        requiresSignature: !!(d.requiresSignature ?? d.toba_requiressignature),
        driveId: d.driveId || d.toba_driveid || undefined,
        itemId: d.itemId || d.toba_itemid || undefined,
        source: d.source || d.toba_source || ((d.driveId || d.toba_driveid) ? 'sharepoint' : undefined)
      }));
      const selectedUsers = (recs || []).map((r: any) => ({ id: r.email || r.user || r.userPrincipalName || r.id || r.email, displayName: r.displayName || r.email, userPrincipalName: r.email, department: r.department, jobTitle: r.jobTitle } as any));
      const selectedGroups: GraphGroup[] = [];
      // Map user -> business
      const nextMap: Record<string, number | null> = {};
      for (const r of recs) {
        const emailLower = String(r.email || r.user || '').toLowerCase();
        if (emailLower) nextMap[emailLower] = r.businessId != null ? Number(r.businessId) : null;
      }
      // Track originals for diffing
      setOriginalRecipientEmails(new Set((recs || []).map((r: any) => String(r.email || r.user || '').trim().toLowerCase()).filter(Boolean)));
      setOriginalDocUrls(new Set((docs || []).map((d: any) => String(d.url || d.webUrl || '').trim()).filter(Boolean)));

      setBatchForm({
        name: String(b.toba_name || b.name || ''),
        startDate: (b.toba_startdate || b.startDate || '') || '',
        dueDate: (b.toba_duedate || b.dueDate || '') || '',
        description: String(b.description || ''),
        selectedUsers: selectedUsers as any,
        selectedGroups,
        selectedDocuments,
        notifyByEmail: true,
        notifyByTeams: false
      });
      setBusinessMap(nextMap);
      setDefaultBusinessId('');
      setEditingBatchId(String(id));
      setActiveTab('batch');
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Loaded batch into editor' } }));
    } catch (e) {
      console.error('Failed to load batch for editing', e);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to open batch for editing' } }));
    }
  };

  return (
    <div className="container">
      <div className="card">
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 24, color: 'var(--primary)' }}>Admin Panel</h1>
            <p className="small muted">Role: {role} • {canEditAdmin ? 'Full Access' : 'Read Only'}</p>
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {sqliteEnabled && (
              <div className="small" title="SQLite API health" style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 8px', border: '1px solid #eee', borderRadius: 999 }}>
                <span style={{ width: 10, height: 10, borderRadius: '50%', background: apiHealth==='ok' ? '#28a745' : apiHealth==='down' ? '#dc3545' : '#ffc107' }} />
                <span>API: {apiHealth === 'ok' ? 'OK' : apiHealth === 'down' ? 'Down' : '—'}</span>
                <button className="btn ghost sm" onClick={pingApi} style={{ marginLeft: 6 }}>Refresh</button>
              </div>
            )}
            <button className="btn ghost sm" onClick={() => {
              try {
                // Simple CSV export of selected documents
                const rows = [
                  ['Title','Url'],
                  ...batchForm.selectedDocuments.map(d => [d.title, d.url])
                ];
                const csv = rows.map(r => r.map(x => '"' + String(x).replace(/"/g, '""') + '"').join(',')).join('\n');
                const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = 'sunbeth-export.csv'; a.click();
                URL.revokeObjectURL(url);
              } catch (e) { console.warn('Export failed', e); }
            }}>Export Data</button>
            <button className="btn ghost sm" onClick={async () => {
              setHealthOpen(true);
              setHealthSteps(null);
              try { setHealthSteps(await runAuthAndGraphCheck()); } catch (e) { setHealthSteps([{ name: 'Health check', ok: false, detail: String(e) }]); }
            }}>System Health</button>
            <button 
              className="btn ghost sm" 
              onClick={() => setShowDebugConsole(true)}
              title="Open batch creation debug console"
            >
              🔍 Debug Logs
            </button>
            {sqliteEnabled && canEditAdmin && (
              <button className="btn ghost sm" onClick={async () => {
                try {
                  const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
                  const email = account?.username || 'seed.user@sunbeth.com';
                  const res = await fetch(`${base}/api/seed?email=${encodeURIComponent(email)}`, { method: 'POST' });
                  if (!res.ok) throw new Error('seed_failed');
                  window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Seeded demo data' } }));
                } catch {
                  window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Seed failed' } }));
                }
              }}>Seed Data</button>
            )}
          </div>
        </div>

        {/* Tab Navigation */}
        <div style={{ display: 'flex', gap: 4, marginBottom: 24, borderBottom: '2px solid #f0f0f0' }}>
          {tabs.map(tab => (
            <button 
              key={tab.id}
              className={activeTab === tab.id ? 'btn sm' : 'btn ghost sm'}
              onClick={() => setActiveTab(tab.id as any)}
              style={{ borderRadius: '8px 8px 0 0' }}
            >
              {tab.icon} {tab.label}
            </button>
          ))}
        </div>

        {/* Tab Content */}
        {activeTab === 'overview' && (
          <div>
            <h2 style={{ fontSize: 18, marginBottom: 16 }}>System Overview</h2>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16 }}>
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: 28, fontWeight: 'bold', color: 'var(--primary)' }}>{overviewStats?.activeBatches ?? '—'}</div>
                <div className="small muted">Active Batches</div>
              </div>
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: 28, fontWeight: 'bold', color: '#28a745' }}>{overviewStats?.totalUsers?.toLocaleString?.() ?? '—'}</div>
                <div className="small muted">Total Users</div>
              </div>
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: 28, fontWeight: 'bold', color: '#ffc107' }}>{overviewStats ? `${overviewStats.completionRate}%` : '—'}</div>
                <div className="small muted">Completion Rate</div>
              </div>
              <div className="card" style={{ padding: 16, textAlign: 'center' }}>
                <div style={{ fontSize: 28, fontWeight: 'bold', color: '#17a2b8' }}>{overviewStats?.overdueBatches ?? 0}</div>
                <div className="small muted">Overdue Batches</div>
              </div>
            </div>

            {/* Permissions Status */}
            <div className="card" style={{ marginTop: 16, padding: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <div>
                  <div style={{ fontWeight: 700, color: 'var(--primary)' }}>Permissions Status</div>
                  <div className="muted small">Required Microsoft Graph scopes</div>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button className="btn ghost sm" onClick={() => checkPermissions()}>Refresh</button>
                  <button className="btn sm" onClick={async () => {
                    try {
                      setGranting(true);
                      // Request all needed scopes in a user-friendly sequence
                      for (const s of requiredScopes) { try { await getGraphToken([s]); } catch {} }
                      await checkPermissions();
                      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Permission prompts completed' } }));
                    } finally { setGranting(false); }
                  }} disabled={granting}>Grant All</button>
                </div>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: 8 }}>
                {requiredScopes.map(s => (
                  <div key={s} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, border: '1px solid #eee', borderRadius: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: permStatus[s] ? '#28a745' : '#dc3545' }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 600 }}>{s}</div>
                      {!permStatus[s] && (
                        <button className="btn ghost sm" onClick={async () => { try { await getGraphToken([s]); } catch {}; await checkPermissions(); }}>Grant {s}</button>
                      )}
                      {permStatus[s] && <div className="small muted">Granted</div>}
                    </div>
                  </div>
                ))}
              </div>
            </div>


          </div>
        )}

        {activeTab === 'settings' && <AdminSettings canEdit={canEditAdmin} />}

        {activeTab === 'manage' && (
          <div style={{ display: 'grid', gap: 16 }}>
            <div className="card" style={{ padding: 16 }}>
              <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Batches</h3>
              <div className="small muted" style={{ marginBottom: 8 }}>View, edit, or delete batches. Deleting a batch removes its documents, recipients, and acknowledgements.</div>
              <ManageBatches canEdit={canEditAdmin} onEdit={(id) => startEditBatch(id)} />
            </div>
            <div className="card" style={{ padding: 16 }}>
              <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Businesses</h3>
              <div className="small muted" style={{ marginBottom: 8 }}>Create, edit, or delete businesses. Deleting a business will unassign it from any recipients mapped to it.</div>
              <BusinessesManager canEdit={canEditAdmin} />
            </div>
          </div>
        )}

        {activeTab === 'batch' && (
          <div>
            <h2 style={{ fontSize: 18, marginBottom: 16 }}>{editingBatchId ? 'Edit Batch' : 'Create New Batch'}</h2>
            
            {/* Batch Details */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 24 }}>
              <div>
                <label className="small">Batch Name:</label>
                <input 
                  type="text" 
                  value={batchForm.name} 
                  onChange={e => setBatchForm({...batchForm, name: e.target.value})}
                  placeholder="Q1 2025 - Code of Conduct"
                  style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
                />
              </div>
              <div>
                <label className="small">Description:</label>
                <input 
                  type="text" 
                  value={batchForm.description} 
                  onChange={e => setBatchForm({...batchForm, description: e.target.value})}
                  placeholder="Annual policy acknowledgement"
                  style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
                />
              </div>
              <div>
                <label className="small">Start Date:</label>
                <input 
                  type="date" 
                  value={batchForm.startDate} 
                  onChange={e => setBatchForm({...batchForm, startDate: e.target.value})}
                  style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
                />
              </div>
              <div>
                <label className="small">Due Date:</label>
                <input 
                  type="date" 
                  value={batchForm.dueDate} 
                  onChange={e => setBatchForm({...batchForm, dueDate: e.target.value})}
                  style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
                />
              </div>
            </div>

            {/* Assignment Section */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <div className="small muted">Choose how you want to select recipients and documents.</div>
              <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="checkbox" checked={useModalSelectors} onChange={e => { setUseModalSelectors(e.target.checked); try { localStorage.setItem(MODAL_TOGGLE_KEY, e.target.checked ? 'true' : 'false'); } catch {} }} />
                Use modal selectors
              </label>
            </div>

            {!useModalSelectors ? (
              <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 24, marginBottom: 24 }}>
                <UserGroupSelector onSelectionChange={(selection) => setBatchForm({...batchForm, selectedUsers: selection.users, selectedGroups: selection.groups})} />
                <SharePointBrowser onDocumentSelect={(spDocs) => setBatchForm({
                  ...batchForm,
                  selectedDocuments: spDocs.map(d => ({
                    title: d.name,
                    url: d.webUrl,
                    version: 1,
                    requiresSignature: false,
                    driveId: (d as any)?.parentReference?.driveId,
                    itemId: (d as any)?.id,
                    source: 'sharepoint'
                  }))
                })} />
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
                        const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
                        const sel = businessMap[email] ?? '';
                        return (
                          <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, alignItems: 'center' }}>
                            <div>
                              <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</div>
                              <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                            </div>
                            <select value={String(sel)} onChange={e => setUserBusiness(email, e.target.value ? Number(e.target.value) : null)}>
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
                        const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
                        const sel = businessMap[email] ?? '';
                        return (
                          <div key={u.id} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, alignItems: 'center' }}>
                            <div style={{ overflow: 'hidden' }}>
                              <div style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</div>
                              <div className="small muted" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{email}</div>
                            </div>
                            <select value={String(sel)} onChange={e => setUserBusiness(email, e.target.value ? Number(e.target.value) : null)}>
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
                  <input type="checkbox" checked={batchForm.notifyByEmail} onChange={e => setBatchForm({...batchForm, notifyByEmail: e.target.checked})} />
                  <span className="small">Email notification (Microsoft Graph)</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: 8, opacity: .6 }} title="Requires Teams Chat.ReadWrite; coming soon">
                  <input type="checkbox" checked={batchForm.notifyByTeams} onChange={e => setBatchForm({...batchForm, notifyByTeams: e.target.checked})} disabled />
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
                <button className="btn ghost" onClick={() => { setBatchForm({name: '', startDate: '', dueDate: '', description: '', selectedUsers: [], selectedGroups: [], selectedDocuments: [], notifyByEmail: true, notifyByTeams: false}); setBusinessMap({}); setDefaultBusinessId(''); setEditingBatchId(null); }}>
                  {editingBatchId ? 'Cancel Edit' : 'Reset Form'}
                </button>
                <button className="btn ghost" title="Preview expanded recipients" onClick={async () => {
                  try {
                    const recipientSet = new Set<string>();
                    for (const u of batchForm.selectedUsers) {
                      const addr = (u.mail || u.userPrincipalName || '').trim();
                      if (addr) recipientSet.add(addr.toLowerCase());
                    }
                    if (batchForm.selectedGroups.length > 0) {
                      const token = await getGraphToken(['Group.Read.All','User.Read']);
                      const arrays = await Promise.all(batchForm.selectedGroups.map(g => getGroupMembers(token, g.id).catch(() => [])));
                      const members = ([] as GraphUser[]).concat(...arrays);
                      for (const m of members) {
                        const addr = (m.mail || m.userPrincipalName || '').trim();
                        if (addr) recipientSet.add(addr.toLowerCase());
                      }
                    }
                    const count = recipientSet.size;
                    window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Recipient preview: ${count} unique addresses` } }));
                  } catch (e) {
                    window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to preview recipients' } }));
                  }
                }}>Preview Recipients</button>
                <button className="btn ghost" title="Grant Graph permissions" onClick={async () => {
                  try {
                    // Trigger consent prompts for common scopes used in Admin
                    await getGraphToken(['User.Read.All','Group.Read.All']);
                    await getGraphToken(['Mail.Send']);
                    await getGraphToken(['Sites.Read.All','Files.ReadWrite.All']);
                    window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Permissions granted (if consented)' } }));
                  } catch (e) {
                    window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Permission grant failed' } }));
                  }
                }}>Grant Permissions</button>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'analytics' && <AnalyticsDashboard />}
      </div>

      {/* Health Modal */}
      {/* Selectors Modals */}
      {useModalSelectors && (
        <>
          <Modal open={usersModalOpen} onClose={() => setUsersModalOpen(false)} title="Assign to Users & Groups" width={800}>
            <UserGroupSelector onSelectionChange={(selection) => setBatchForm({...batchForm, selectedUsers: selection.users, selectedGroups: selection.groups})} />
          </Modal>
          <Modal open={docsModalOpen} onClose={() => setDocsModalOpen(false)} title="SharePoint Documents" width={920}>
            <SharePointBrowser onDocumentSelect={(spDocs) => setBatchForm({
              ...batchForm,
              selectedDocuments: spDocs.map(d => ({
                title: d.name,
                url: d.webUrl,
                version: 1,
                requiresSignature: false,
                driveId: (d as any)?.parentReference?.driveId,
                itemId: (d as any)?.id,
                source: 'sharepoint'
              }))
            })} />
          </Modal>
        </>
      )}

      {healthOpen && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div className="card" style={{ width: 520, maxWidth: '90%', padding: 16 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <h3 style={{ margin: 0 }}>System Health</h3>
              <button className="btn ghost sm" onClick={() => setHealthOpen(false)}>Close</button>
            </div>
            {!healthSteps && <div className="small muted">Running checks...</div>}
            {healthSteps && (
              <div style={{ display: 'grid', gap: 8 }}>
                {healthSteps.map((s, i) => (
                  <div key={i} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: 8, border: '1px solid #eee', borderRadius: 6 }}>
                    <span style={{ width: 10, height: 10, borderRadius: '50%', background: s.ok ? '#28a745' : '#dc3545' }} />
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 600 }}>{s.name}</div>
                      {s.detail && <div className="small muted" style={{ whiteSpace: 'pre-wrap' }}>{s.detail}</div>}
                    </div>
                  </div>
                ))}
                {/* Permissions quick-fix */}
                <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                  <button className="btn ghost sm" onClick={async () => {
                    try {
                      setGranting(true);
                      // Grant core scopes in series
                      await getGraphToken(['User.Read']);
                      await getGraphToken(['User.Read.All']);
                      await getGraphToken(['Group.Read.All']);
                      await getGraphToken(['Sites.Read.All']);
                      await getGraphToken(['Files.ReadWrite.All']);
                      await getGraphToken(['Mail.Send']);
                      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Requested missing permissions' } }));
                    } finally {
                      setGranting(false);
                      try { setHealthSteps(await runAuthAndGraphCheck()); } catch {}
                    }
                  }} disabled={granting}>Grant Missing Permissions</button>
                  <button className="btn ghost sm" onClick={async () => { try { setHealthSteps(await runAuthAndGraphCheck()); } catch {} }}>Re-run</button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
      
      {/* Batch Creation Debug Console */}
      <BatchCreationDebug 
        isVisible={showDebugConsole}
        onClose={() => setShowDebugConsole(false)}
      />
    </div>
  );
};

export default AdminPanel;

// --- Admin helpers: Businesses & Batches managers ---
type Biz = { id: number; name: string; code?: string; isActive?: boolean; description?: string };
const apiBase = () => ((process.env.REACT_APP_API_BASE as string) || '').replace(/\/$/, '');
const sqliteOn = () => (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;

const BusinessesManager: React.FC<{ canEdit: boolean }> = ({ canEdit }) => {
  const [items, setItems] = useState<Biz[]>([]);
  const [busy, setBusy] = useState(false);
  const [form, setForm] = useState<{ name: string; code: string; isActive: boolean; description: string }>({ name: '', code: '', isActive: true, description: '' });
  const [editRow, setEditRow] = useState<Record<number, Partial<Biz>>>({});

  const load = async () => {
    if (!sqliteOn()) return;
    try { const res = await fetch(`${apiBase()}/api/businesses`); const j = await res.json(); setItems(Array.isArray(j) ? j : []); } catch { setItems([]); }
  };
  useEffect(() => { load(); }, []);

  const create = async () => {
    if (!canEdit || !sqliteOn()) return;
    const name = form.name.trim(); if (!name) { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enter a business name' } })); return; }
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
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Business created' } }));
    } catch { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to create business' } })); }
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
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Business updated' } }));
    } catch { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to update business' } })); }
    finally { setBusy(false); }
  };

  const del = async (id: number) => {
    if (!canEdit || !sqliteOn()) return;
    if (!confirm('Delete this business? This will unassign it from any recipients.')) return;
    setBusy(true);
    try {
      await deleteBusiness(id);
      await load();
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Business deleted' } }));
    } catch { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to delete business' } })); }
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
          const row = editRow[b.id] || {};
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

const ManageBatches: React.FC<{ canEdit: boolean; onEdit: (id: string) => void }> = ({ canEdit, onEdit }) => {
  const [items, setItems] = useState<Array<{ toba_batchid: string; toba_name: string; toba_startdate?: string; toba_duedate?: string; toba_status?: string }>>([]);
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
    if (!confirm('Delete this batch and all related records?')) return;
    setBusy(true);
    try {
      const res = await fetch(`${apiBase()}/api/batches/${encodeURIComponent(id)}`, { method: 'DELETE' });
      if (!res.ok) throw new Error('delete_failed');
      await load();
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Batch deleted' } }));
    } catch {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to delete batch' } }));
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
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Batch updated' } }));
    } catch {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to update batch' } }));
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
