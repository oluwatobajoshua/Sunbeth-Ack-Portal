import React, { useEffect, useState } from 'react';
import { useRBAC } from '../context/RBACContext';
// import { useAuth } from '../context/AuthContext';
import { GraphUser, GraphGroup, getUsers, getGroups, getOrganizationStructure, UserSearchFilters, getGroupMembers } from '../services/graphUserService';
import { SharePointSite, SharePointDocumentLibrary, SharePointDocument, getSharePointSites, getDocumentLibraries, getDocuments, uploadFileToDrive, getFolderItems } from '../services/sharepointService';
import { useRuntimeMock } from '../utils/runtimeMock';
import AnalyticsDashboard from './AnalyticsDashboard';
import Modal from './Modal';
import { sendEmail, buildBatchEmail /*, sendTeamsDirectMessage*/ } from '../services/notificationService';
import { getGraphToken } from '../services/authTokens';
// logger import removed (Dataverse logging no longer used)
import { runAuthAndGraphCheck, Step } from '../diagnostics/health';
// Business type import removed (Dataverse features removed)
import { provisionSharePointLists, type ProvisionStep as SPProvisionStep, spCreateBatch, spCreateDocument, spCreateRecipient, findSitesByName, setSharePointSiteIdOverride } from '../services/spListsService';

// Enhanced Admin Settings Component
type AdminSettingsProps = { canEdit: boolean };

const AdminSettings: React.FC<AdminSettingsProps> = ({ canEdit }) => {
  const { account } = useAuthCtx();
  const storageKey = 'mock_admin_settings';
  const [settings, setSettings] = useState({
    enableUpload: false,
    requireSig: false,
    autoReminder: true,
    reminderDays: 3,
    allowBulkAssignment: true,
    requireApproval: false
  });
  // Dataverse functionality removed
  const [spProvisioning, setSpProvisioning] = useState(false);
  const [spProvisionLogs, setSpProvisionLogs] = useState<SPProvisionStep[] | null>(null);
  const [spSiteQuery, setSpSiteQuery] = useState<string>('Sunbeth Intranet');
  const [spSiteMatches, setSpSiteMatches] = useState<Array<{ id: string; displayName: string; webUrl: string }>>([]);

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

  // Dataverse provisioning/seed removed

  const provisionSharePoint = async () => {
    try {
      // Allow provisioning if we have a site id from env or override
      const overrideId = (() => { try { return localStorage.getItem('sunbeth:sp:siteIdOverride'); } catch { return null; } })();
      if (!process.env.REACT_APP_SP_SITE_ID && !overrideId) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Choose a SharePoint site first (Use Site), or set REACT_APP_SP_SITE_ID.' } }));
        return;
      }
      setSpProvisioning(true);
      setSpProvisionLogs(null);
      // Ensure we have Sites.ReadWrite.All before provisioning
      try { await getGraphToken(['Sites.ReadWrite.All']); } catch {}
      const logs = await provisionSharePointLists();
      setSpProvisionLogs(logs);
      const ok = logs.every(l => l.ok);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: ok ? 'SharePoint lists provisioned' : 'SP provision completed with issues (check details)' } }));
      try { console.table(logs); } catch {}
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'SharePoint provisioning failed' } }));
    } finally { setSpProvisioning(false); }
  };

  const resolveAndUseSharePointSite = async () => {
    try {
      // Find sites matching the query
      const results = await findSitesByName(spSiteQuery);
      setSpSiteMatches(results);
      if (!results.length) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'No SharePoint sites found with that name.' } }));
        return;
      }
      // Prefer exact display name match (case-insensitive)
      const exact = results.find(s => (s.displayName || '').toLowerCase() === spSiteQuery.trim().toLowerCase());
      const chosen = exact || results[0];
      setSharePointSiteIdOverride(chosen.id);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Using site: ${chosen.displayName || chosen.webUrl}` } }));
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to find SharePoint site' } }));
    }
  };

  // Dataverse write test removed

  // Dataverse test removed

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

  // Dataverse read test removed

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
        {canEdit && <button className="btn ghost" onClick={provisionSharePoint} title="Creates SharePoint Lists and columns via Microsoft Graph" disabled={spProvisioning}>
          {spProvisioning ? 'Provisioning…' : 'Provision SharePoint Lists'}
        </button>}
        {canEdit && <button className="btn ghost" onClick={seedSqliteForMe} title="Seed SQLite with a demo batch, docs, and recipients for your account">Seed SQLite (for me)</button>}
        {/* SharePoint site selection */}
        {canEdit && (
          <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
            <input
              type="text"
              value={spSiteQuery}
              onChange={e => setSpSiteQuery(e.target.value)}
              placeholder="SharePoint site name (e.g., Sunbeth Intranet)"
              style={{ padding: 6, border: '1px solid #ddd', borderRadius: 4, minWidth: 240 }}
            />
            <button className="btn ghost" onClick={resolveAndUseSharePointSite} title="Find and use this SharePoint site for provisioning">Use Site</button>
          </div>
        )}
  {canEdit && <button className="btn ghost" onClick={grantCorePermissions} title="Request common Microsoft Graph permissions in one go">Grant Core Permissions</button>}
      </div>

      {/* Inline provisioning results for immediate feedback */}
      {/* Dataverse provisioning logs removed */}
      {spProvisionLogs && (
        <div className="card" style={{ marginTop: 8, padding: 12 }}>
          <div style={{ fontWeight: 700, color: 'var(--primary)', marginBottom: 6 }}>SharePoint Provisioning Results</div>
          <div className="small muted" style={{ marginBottom: 8 }}>These steps were executed against your SharePoint site via Microsoft Graph.</div>
          <div style={{ display: 'grid', gap: 6 }}>
            {spProvisionLogs.map((log, idx) => (
              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, border: '1px solid #eee', borderRadius: 6, padding: 8 }}>
                <span style={{ width: 10, height: 10, borderRadius: '50%', background: log.ok ? '#28a745' : '#dc3545' }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600 }}>{log.step}</div>
                  {log.detail && <div className="small muted" style={{ wordBreak: 'break-word' }}>{log.detail}</div>}
                </div>
              </div>
            ))}
          </div>
          {spSiteMatches && spSiteMatches.length > 0 && (
            <div style={{ marginTop: 12 }}>
              <div className="small" style={{ fontWeight: 600, marginBottom: 6 }}>Matched Sites</div>
              <div style={{ display: 'grid', gap: 6 }}>
                {spSiteMatches.map(s => (
                  <div key={s.id} style={{ display: 'flex', gap: 8, alignItems: 'center', border: '1px dashed #eee', padding: 8, borderRadius: 6 }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 600 }}>{s.displayName || 'Site'}</div>
                      <div className="small muted" style={{ wordBreak: 'break-word' }}>{s.webUrl}</div>
                      <div className="small muted" style={{ wordBreak: 'break-word' }}>{s.id}</div>
                    </div>
                    <button className="btn ghost sm" onClick={() => { setSharePointSiteIdOverride(s.id); window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Using site: ${s.displayName || s.webUrl}` } })); }}>Use</button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// User/Group Selection Component
const UserGroupSelector: React.FC<{ onSelectionChange: (selection: any) => void }> = ({ onSelectionChange }) => {
  const { getToken, login, account } = useAuthCtx();
  const runtimeMock = useRuntimeMock();
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
    if (runtimeMock) {
      // Mock data for testing
      setUsers([
        { id: '1', displayName: 'John Smith', userPrincipalName: 'john@sunbeth.com', department: 'HR', jobTitle: 'HR Manager' },
        { id: '2', displayName: 'Jane Doe', userPrincipalName: 'jane@sunbeth.com', department: 'IT', jobTitle: 'Developer' },
        { id: '3', displayName: 'Bob Wilson', userPrincipalName: 'bob@sunbeth.com', department: 'Finance', jobTitle: 'Accountant' }
      ]);
      setGroups([
        { id: 'g1', displayName: 'All Employees', groupTypes: [], memberCount: 150 },
        { id: 'g2', displayName: 'HR Team', groupTypes: [], memberCount: 5 },
        { id: 'g3', displayName: 'IT Department', groupTypes: [], memberCount: 12 }
      ]);
      setOrgStructure({ departments: ['HR', 'IT', 'Finance'], jobTitles: ['Manager', 'Developer', 'Accountant'], locations: ['New York', 'London'] });
      return;
    }

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
  }, [filters, runtimeMock]);

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
      {!runtimeMock && (
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
      )}
      
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

// SharePoint Document Browser Component
import { useAuth as useAuthCtx } from '../context/AuthContext';
const SharePointBrowser: React.FC<{ onDocumentSelect: (docs: SharePointDocument[]) => void }> = ({ onDocumentSelect }) => {
  const { getToken, login, account } = useAuthCtx();
  const runtimeMock = useRuntimeMock();
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
  const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB logical limit (change per policy)

  const loadSites = async () => {
    if (runtimeMock) {
      setSites([
        { id: 'site1', displayName: 'HR Policies Site', webUrl: 'https://sunbeth.sharepoint.com/sites/hr' },
        { id: 'site2', displayName: 'Company Documents', webUrl: 'https://sunbeth.sharepoint.com/sites/docs' }
      ]);
      return;
    }

    setLoading(true);
    setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const sitesData = await getSharePointSites(token);
      setSites(sitesData);
    } catch (error) {
      console.error('Failed to load SharePoint sites:', error);
      setSpError('Failed to load SharePoint sites. Ensure you are signed in and have granted Sites.Read.All and Files.Read.All.');
    } finally {
      setLoading(false);
    }
  };

  const loadLibraries = async (siteId: string) => {
    if (runtimeMock) {
      setLibraries([
        { id: 'lib1', name: 'Documents', displayName: 'Documents', webUrl: 'https://sunbeth.sharepoint.com/sites/hr/Documents', driveType: 'documentLibrary' },
        { id: 'lib2', name: 'Policies', displayName: 'HR Policies', webUrl: 'https://sunbeth.sharepoint.com/sites/hr/Policies', driveType: 'documentLibrary' }
      ]);
      return;
    }

    setLoading(true);
    setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const librariesData = await getDocumentLibraries(token, siteId);
      setLibraries(librariesData);
    } catch (error) {
      console.error('Failed to load document libraries:', error);
      const msg = typeof (error as any)?.message === 'string' ? (error as any).message : '';
      setSpError(`Failed to load document libraries.${msg ? ' ' + msg : ''}`);
    } finally {
      setLoading(false);
    }
  };

  const loadDocuments = async (driveId: string, folderId: string = 'root') => {
    if (runtimeMock) {
      setDocuments([
        { 
          id: 'doc1', 
          name: 'Code of Conduct.pdf', 
          webUrl: 'https://sunbeth.sharepoint.com/sites/hr/Documents/Code%20of%20Conduct.pdf',
          size: 1024000,
          createdDateTime: '2025-01-01T00:00:00Z',
          lastModifiedDateTime: '2025-01-15T00:00:00Z',
          file: { mimeType: 'application/pdf' }
        },
        { 
          id: 'doc2', 
          name: 'Health and Safety Policy.docx', 
          webUrl: 'https://sunbeth.sharepoint.com/sites/hr/Documents/Health%20and%20Safety.docx',
          size: 512000,
          createdDateTime: '2024-12-01T00:00:00Z',
          lastModifiedDateTime: '2025-01-10T00:00:00Z',
          file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
        }
      ]);
      return;
    }

    setLoading(true);
    setSpError(null);
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const documentsData = await getDocuments(token, driveId, folderId, searchQuery);
      setDocuments(documentsData);
    } catch (error) {
      console.error('Failed to load documents:', error);
      setSpError('Failed to load documents.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadSites();
  }, [runtimeMock]);

  useEffect(() => {
    if (selectedSite) {
      loadLibraries(selectedSite);
      setSelectedLibrary('');
      setDocuments([]);
    }
  }, [selectedSite]);

  useEffect(() => {
    if (selectedLibrary) {
  loadDocuments(selectedLibrary, 'root');
      // Also load root folder items for upload picker
      (async () => {
        if (runtimeMock) {
          setFolderItems([{ id: 'root', name: 'Root', folder: { childCount: 0 } }]);
          setSelectedFolderId('root');
          setBreadcrumbs([{ id: 'root', name: 'Root' }]);
          return;
        }
        try {
          const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
          if (!token) throw new Error('No token available');
          const items = await getFolderItems(token, selectedLibrary, 'root');
          setFolderItems(items);
          setSelectedFolderId('root');
          setBreadcrumbs([{ id: 'root', name: 'Root' }]);
        } catch (e) {
          console.error('Failed to load folder items', e);
        }
      })();
    }
  }, [selectedLibrary, searchQuery]);

  // Refresh document listing when navigating folders (browse mode)
  useEffect(() => {
    if (!selectedLibrary) return;
    if (spTab !== 'browse') return;
    loadDocuments(selectedLibrary, selectedFolderId || 'root');
  }, [selectedFolderId, spTab]);

  useEffect(() => {
    const selected = Array.from(selectedDocs).map(id => documents.find(d => d.id === id)!).filter(Boolean);
    onDocumentSelect(selected);
  }, [selectedDocs, documents]);

  const toggleDocument = (docId: string) => {
    const newSelection = new Set(selectedDocs);
    if (newSelection.has(docId)) {
      newSelection.delete(docId);
    } else {
      newSelection.add(docId);
    }
    setSelectedDocs(newSelection);
  };

  const canUpload = !!selectedLibrary;

  const navigateFolder = async (folderId: string, folderName: string) => {
    if (runtimeMock) {
      setSelectedFolderId(folderId);
      setBreadcrumbs(prev => [...prev, { id: folderId, name: folderName }]);
      // mock: no deeper items
      return;
    }
    try {
      const token = await getToken(['Sites.Read.All', 'Files.Read.All']);
      if (!token) throw new Error('No token available');
      const items = await getFolderItems(token, selectedLibrary, folderId);
      setFolderItems(items);
      setSelectedFolderId(folderId);
      setBreadcrumbs(prev => {
        const idx = prev.findIndex(b => b.id === folderId);
        if (idx >= 0) return prev.slice(0, idx + 1);
        return [...prev, { id: folderId, name: folderName }];
      });
    } catch (e) {
      console.error('Failed to navigate folder', e);
    }
  };

  const handleUpload = async (files: FileList | File[] | null) => {
    const arr: File[] = files ? Array.from(files as any) : [];
    if (!arr.length || !selectedLibrary) return;
    setUploading(true);
    setUploadProgress(0);
    setUploadStatuses([]);
    try {
      const token = await getToken(['Files.ReadWrite.All', 'Sites.Read.All']);
      if (!token) throw new Error('No token available');
      const uploadedDocs: SharePointDocument[] = [];
      for (let i = 0; i < arr.length; i++) {
        const f = arr[i]!;
        // Validate max size (optional UX guard; Graph chunked allows much larger within tenant limit)
        if ((f as any).size > MAX_FILE_SIZE) {
          setUploadStatuses(prev => [...prev, { name: f.name, progress: 0, error: `File exceeds ${Math.round(MAX_FILE_SIZE/1024/1024)}MB limit` }]);
          continue;
        }
        setUploadStatuses(prev => [...prev, { name: f.name, progress: 0 }]);
        try {
          const doc = await uploadFileToDrive(token, selectedLibrary, f, f.name, undefined, (p) => {
            setUploadProgress(p);
            setUploadStatuses(prev => {
              const copy = [...prev];
              const idx = copy.findIndex(u => u.name === f.name);
              if (idx >= 0) copy[idx] = { ...copy[idx], progress: p };
              return copy;
            });
          }, selectedFolderId);
          uploadedDocs.push(doc);
          // add to list and select
          setDocuments(prev => [{
            id: doc.id,
            name: doc.name,
            webUrl: doc.webUrl,
            size: (doc as any).size || f.size,
            createdDateTime: (doc as any).createdDateTime || new Date().toISOString(),
            lastModifiedDateTime: (doc as any).lastModifiedDateTime || new Date().toISOString(),
            file: (doc as any).file || { mimeType: (f as any).type || 'application/octet-stream' },
            parentReference: (doc as any).parentReference
          } as SharePointDocument, ...prev]);
          setSelectedDocs(prev => new Set(prev).add(doc.id));
        } catch (err: any) {
          const msg = typeof err?.message === 'string' ? err.message : 'Upload failed';
          setUploadStatuses(prev => {
            const copy = [...prev];
            const idx = copy.findIndex(u => u.name === f.name);
            if (idx >= 0) copy[idx] = { ...copy[idx], error: msg };
            return copy;
          });
        }
      }
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Uploaded ${uploadedDocs.length} file(s)` } }));
      // Switch back to browse to show the files
      setSpTab('browse');
    } catch (e) {
      console.error('Upload failed', e);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Upload failed' } }));
    } finally {
      setUploading(false);
      setUploadProgress(null);
    }
  };

  const onDropFiles = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      void handleUpload(files);
    }
  };

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 16px 0', fontSize: 16 }}>SharePoint Documents</h3>
      {!runtimeMock && (
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
      )}

      {/* Mode Tabs */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, borderBottom: '1px solid #e0e0e0' }}>
        {(['browse', 'upload'] as const).map(t => (
          <button
            key={t}
            className={spTab === t ? 'btn sm' : 'btn ghost sm'}
            onClick={() => setSpTab(t)}
            // Allow navigating to Upload to show guidance even before a library is selected
            // File input remains disabled until a library is chosen.
          >
            {t === 'browse' ? 'Browse' : 'Upload'}
          </button>
        ))}
      </div>
      
      {loading && <div className="small muted">Loading...</div>}

      {/* Site Selection */}
      <div style={{ marginBottom: 16 }}>
        <label className="small">SharePoint Site:</label>
        <select 
          value={selectedSite} 
          onChange={e => setSelectedSite(e.target.value)}
          style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
        >
          <option value="">Select a site...</option>
          {sites.map(site => <option key={site.id} value={site.id}>{site.displayName}</option>)}
        </select>
      </div>

      {/* Library Selection */}
      {selectedSite && (
        <div style={{ marginBottom: 16 }}>
          <label className="small">Document Library:</label>
          <select 
            value={selectedLibrary} 
            onChange={e => setSelectedLibrary(e.target.value)}
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
          >
            <option value="">Select a library...</option>
            {libraries.map(lib => <option key={lib.id} value={lib.id}>{lib.displayName}</option>)}
          </select>
        </div>
      )}

      {/* Search (Browse Mode) */}
      {selectedLibrary && spTab === 'browse' && (
        <div style={{ marginBottom: 16 }}>
          <input 
            type="text" 
            placeholder="Search documents..." 
            value={searchQuery} 
            onChange={e => setSearchQuery(e.target.value)}
            style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4 }}
          />
        </div>
      )}

      {/* Document List (Browse Mode) */}
      {spTab === 'browse' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          {/* Breadcrumbs for folder navigation */}
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
          {/* Folder items (click to navigate) */}
          {selectedLibrary && folderItems.filter(i => i.folder).map(f => (
            <div key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 8px', borderBottom: '1px solid #f5f5f5' }}>
              <button className="btn ghost sm" onClick={() => navigateFolder(f.id, f.name)}>
                📁 {f.name}
              </button>
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
                  <input 
                    type="checkbox" 
                    checked={selectedDocs.has(doc.id)} 
                    onChange={() => toggleDocument(doc.id)} 
                  />
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 500 }}>{doc.name}</div>
                    <div className="small muted">
                      {(doc.size / 1024).toFixed(1)} KB • Modified {new Date(doc.lastModifiedDateTime).toLocaleDateString()}
                    </div>
                    <a href={doc.webUrl} target="_blank" rel="noopener noreferrer" className="small" style={{ color: '#0066cc' }}>
                      View in SharePoint ↗
                    </a>
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
              <select 
                value={selectedFolderId}
                onChange={e => setSelectedFolderId(e.target.value)}
                style={{ width: '100%', padding: 8, border: '1px solid #ddd', borderRadius: 4, marginTop: 4 }}
              >
                <option value="root">/ (root)</option>
                {folderItems.filter(i => i.folder).map(f => (
                  <option key={f.id} value={f.id}>/ {f.name}</option>
                ))}
              </select>
            </div>
          )}
          {/* Drop Zone */}
          <div
            onDragEnter={e => { e.preventDefault(); e.stopPropagation(); if (selectedLibrary && !uploading) setIsDragging(true); }}
            onDragOver={e => { e.preventDefault(); e.stopPropagation(); }}
            onDragLeave={e => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); }}
            onDrop={onDropFiles}
            style={{
              border: '2px dashed ' + (isDragging ? 'var(--primary)' : '#ccc'),
              background: isDragging ? 'rgba(0,0,0,0.02)' : 'transparent',
              padding: 16,
              borderRadius: 8,
              textAlign: 'center',
              opacity: (!selectedLibrary || uploading) ? 0.6 : 1,
              pointerEvents: (!selectedLibrary || uploading) ? 'none' : 'auto'
            }}
          >
            <div className="small" style={{ marginBottom: 8 }}>
              Drag and drop files here
            </div>
            <div className="small muted">or</div>
            <div style={{ marginTop: 8 }}>
              <label className="btn ghost sm" style={{ cursor: (!selectedLibrary || uploading) ? 'not-allowed' : 'pointer' }}>
                Browse files
                <input
                  type="file"
                  multiple
                  disabled={!selectedLibrary || uploading}
                  accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.html"
                  onChange={e => handleUpload(e.target.files)}
                  style={{ display: 'none' }}
                />
              </label>
            </div>
            <div className="small muted" style={{ marginTop: 8 }}>
              Allowed: PDF, Word, Excel, PowerPoint, Text, HTML
            </div>
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
                    {u.error ? (
                      <span style={{ color: '#d33' }}>{u.error}</span>
                    ) : (
                      <span>{u.progress}%</span>
                    )}
                  </div>
                  {!u.error && (
                    <div className="progressBar" aria-hidden="true">
                      <i style={{ width: `${u.progress}%` }} />
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
          <div className="small muted" style={{ marginTop: 8 }}>
            Files are uploaded into the selected folder. Large files use a chunked upload session.
          </div>
        </div>
      )}
    </div>
  );
};

// Main Admin Panel Component
const AdminPanel: React.FC = () => {
  const { role, canSeeAdmin, canEditAdmin } = useRBAC();
  const runtimeMock = useRuntimeMock();
  const [activeTab, setActiveTab] = useState<'overview' | 'settings' | 'batch' | 'analytics'>('overview');
  const [healthOpen, setHealthOpen] = useState(false);
  const [healthSteps, setHealthSteps] = useState<Step[] | null>(null);
  const [granting, setGranting] = useState(false);
  const [permStatus, setPermStatus] = useState<Record<string, boolean>>({});
  // Dataverse state removed
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
  const [usersModalOpen, setUsersModalOpen] = useState(false);
  const [docsModalOpen, setDocsModalOpen] = useState(false);

  const requiredScopes = ['User.Read','User.Read.All','Group.Read.All','Sites.Read.All','Files.ReadWrite.All','Mail.Send'];

  const checkPermissions = async () => {
    const status: Record<string, boolean> = {};
    for (const scope of requiredScopes) {
      try { await getGraphToken([scope]); status[scope] = true; }
      catch { status[scope] = false; }
    }
    setPermStatus(status);
  };

  useEffect(() => { if (!adminLight) { checkPermissions().catch(() => {}); } }, [adminLight]);
  // Dataverse whoAmI/status check removed
  const [batchForm, setBatchForm] = useState<{
    name: string;
    startDate: string;
    dueDate: string;
    description: string;
    selectedUsers: GraphUser[];
    selectedGroups: GraphGroup[];
    selectedDocuments: SharePointDocument[];
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

  // Dataverse business assignment state removed

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

  // Dataverse businesses loader removed

  const tabs = [
    { id: 'overview', label: 'Overview', icon: '📊' },
    { id: 'settings', label: 'Settings', icon: '⚙️' },
    { id: 'batch', label: 'Create Batch', icon: '📝' },
    { id: 'analytics', label: 'Analytics', icon: '📈' }
  ];
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  const [overviewStats, setOverviewStats] = useState<{ totalBatches: number; activeBatches: number; totalUsers: number; completionRate: number; overdueBatches: number; avgCompletionTime: number } | null>(null);
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

  const createBatch = async () => {
    try {
      // Dataverse path removed; use SharePoint Lists only if enabled
      const spEnabled = process.env.REACT_APP_ENABLE_SP_LISTS === 'true';
      let spBatchId: number | null | undefined;
      // SP Lists path (optional, independent of DV)
      if (spEnabled && process.env.REACT_APP_SP_SITE_ID) {
        try {
          spBatchId = await spCreateBatch({ name: batchForm.name, startDate: batchForm.startDate || undefined, dueDate: batchForm.dueDate || undefined, description: batchForm.description || undefined, status: 1 });
        } catch (e) {
          console.warn('SharePoint batch creation error', e);
        }
      }

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

      if (batchForm.notifyByEmail && recipients.length > 0) {
        await sendEmail(recipients, subject, bodyHtml);
      }
      // Teams optional (requires Chat.ReadWrite)
      // if (batchForm.notifyByTeams) {
      //   const userIds = batchForm.selectedUsers.map(u => u.id);
      //   await sendTeamsDirectMessage(userIds, `New acknowledgement assigned: ${batchForm.name}`);
      // }

      // Dataverse document persistence removed
      // 2b) Persist documents in SharePoint Lists (if enabled)
      if (spEnabled && spBatchId) {
        for (const d of batchForm.selectedDocuments) {
          try {
            await spCreateDocument(spBatchId, { title: d.name, url: d.webUrl, version: 1, requiresSignature: false });
          } catch (e) {
            console.warn('SP Lists document creation error', d.name, e);
          }
        }
      }

      // Dataverse recipient persistence removed
      // 3b) Persist Recipients in SharePoint Lists (if enabled)
      if (spEnabled && spBatchId) {
        const chooseBusinessSp = (emailLower: string): number | null => {
          // We only have string ids from DV businesses; for SP we haven't modeled businesses link-ups yet.
          // Leave as null for now; business can be added later in SP list if needed.
          return null;
        };
        for (const r of recipients) {
          try {
            const emailLower = (r.address || '').toLowerCase();
            const u = userByEmailLower.get(emailLower);
            let primaryGroupName: string | undefined = undefined;
            const origins = recipientOrigins.get(emailLower);
            if (origins && origins.size > 0) {
              const firstGid = origins.values().next().value as string;
              const g = batchForm.selectedGroups.find(x => x.id === firstGid);
              if (g?.displayName) primaryGroupName = g.displayName;
            }
            const bizId = chooseBusinessSp(emailLower);
            await spCreateRecipient(spBatchId, {
              businessId: bizId,
              user: emailLower,
              email: emailLower,
              displayName: r.name || undefined,
              department: u?.department || undefined,
              jobTitle: u?.jobTitle || undefined,
              location: u?.officeLocation || undefined,
              primaryGroup: primaryGroupName || undefined
            });
          } catch (e) {
            console.warn('SP Lists recipient creation error', r.address, e);
          }
        }
      }

      // Final feedback
      if (spBatchId) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Batch "${batchForm.name}" created in SharePoint${batchForm.notifyByEmail ? ' and email sent' : ''}.` } }));
      } else {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Batch created (notifications sent), but saving to SharePoint failed.` } }));
      }

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
      // Dataverse business selections cleared (state removed)
    } catch (e) {
      console.error('Create batch failed', e);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Failed to create batch or send notifications' } }));
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
          <div style={{ display: 'flex', gap: 8 }}>
            <button className="btn ghost sm" onClick={() => {
              try {
                // Simple CSV export of selected documents
                const rows = [
                  ['Name','Url','Size','Modified'],
                  ...batchForm.selectedDocuments.map(d => [d.name, d.webUrl, String(d.size), new Date(d.lastModifiedDateTime).toISOString()])
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

            {/* Dataverse sections removed */}
          </div>
        )}

        {activeTab === 'settings' && <AdminSettings canEdit={canEditAdmin} />}

        {activeTab === 'batch' && (
          <div>
            <h2 style={{ fontSize: 18, marginBottom: 16 }}>Create New Batch</h2>
            
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
                <SharePointBrowser onDocumentSelect={(docs) => setBatchForm({...batchForm, selectedDocuments: docs})} />
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
                    <div style={{ fontWeight: 700 }}>SharePoint Documents</div>
                    <div className="small muted">{batchForm.selectedDocuments.length} document(s) selected</div>
                  </div>
                  <button className="btn sm" onClick={() => setDocsModalOpen(true)}>Choose documents</button>
                </div>
              </div>
            )}

            {/* Dataverse Business Assignment removed */}

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
                  onClick={createBatch}
                  disabled={!batchForm.name || !batchForm.startDate || !batchForm.dueDate || batchForm.selectedDocuments.length === 0}
                >
                  Create Batch
                </button>
                <button className="btn ghost" onClick={() => setBatchForm({name: '', startDate: '', dueDate: '', description: '', selectedUsers: [], selectedGroups: [], selectedDocuments: [], notifyByEmail: true, notifyByTeams: false})}>
                  Reset Form
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
            <SharePointBrowser onDocumentSelect={(docs) => setBatchForm({...batchForm, selectedDocuments: docs})} />
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
    </div>
  );
};

export default AdminPanel;
