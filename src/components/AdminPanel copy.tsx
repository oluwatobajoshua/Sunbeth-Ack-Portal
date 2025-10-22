import React, { useEffect, useState } from 'react';
import { useRBAC } from '../context/RBACContext';
// import { useAuth } from '../context/AuthContext';
import { GraphUser, GraphGroup, getUsers, getGroups, getOrganizationStructure, UserSearchFilters, getGroupMembers } from '../services/graphUserService';
import { SharePointSite, SharePointDocumentLibrary, SharePointDocument, getSharePointSites, getDocumentLibraries, getDocuments, uploadFileToDrive, getFolderItems } from '../services/sharepointService';
import { useRuntimeMock } from '../utils/runtimeMock';
import AnalyticsDashboard from './AnalyticsDashboard';
import DataverseExplorer from './DataverseExplorer';
import Modal from './Modal';
import { sendEmail, buildBatchEmail /*, sendTeamsDirectMessage*/ } from '../services/notificationService';
import { getDataverseToken, getGraphToken } from '../services/authTokens';
import { info as logInfo } from '../diagnostics/logger';
import { runAuthAndGraphCheck, Step } from '../diagnostics/health';
import { DV_SETS, DV_ATTRS } from '../services/dataverseConfig';
import type { Business } from '../types/models';
import { getBusinesses, countRecords, probeReadAccess } from '../services/dataverseService';
import { provisionSunbethSchema, whoAmI, seedSunbethSampleData, resolveEntitySets, dvWriteTest, type ProvisionLog } from '../services/dataverseProvisioning';
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
  const [provisioning, setProvisioning] = useState(false);
  const [provisionLogs, setProvisionLogs] = useState<ProvisionLog[] | null>(null);
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

  const provision = async () => {
    try {
      const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const url = process.env.REACT_APP_DATAVERSE_URL;
      if (!enabled || !url) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL to provision.' } }));
        return;
      }
      setProvisioning(true);
      setProvisionLogs(null);
      const logs = await provisionSunbethSchema();
      setProvisionLogs(logs);
      const ok = logs.every(l => l.ok);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: ok ? 'Dataverse schema provisioned' : 'Provision completed with issues (check console)' } }));
      try { console.table(logs); } catch {}
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Provisioning failed' } }));
    } finally { setProvisioning(false); }
  };
  const seedSamples = async () => {
    try {
      const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const url = process.env.REACT_APP_DATAVERSE_URL;
      if (!enabled || !url) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL to seed data.' } }));
        return;
      }
      setProvisioning(true);
      setProvisionLogs(null);
      const logs = await seedSunbethSampleData();
      setProvisionLogs(logs);
      const ok = logs.every(l => l.ok);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: ok ? 'Sample data seeded' : 'Seeding completed with issues (check console)' } }));
      try { console.table(logs); } catch {}
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Seeding failed' } }));
    } finally { setProvisioning(false); }
  };

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

  const writeTest = async () => {
    try {
      const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const url = process.env.REACT_APP_DATAVERSE_URL;
      if (!enabled || !url) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL to run DV write test.' } }));
        return;
      }
      setProvisioning(true);
      setProvisionLogs(null);
      const logs = await dvWriteTest();
      setProvisionLogs(logs);
      const ok = logs.every(l => l.ok);
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: ok ? 'DV write test OK (created/fetched/deleted)' : 'DV write test had issues (see details)' } }));
      try { console.table(logs); } catch {}
    } catch (e) {
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'DV write test failed' } }));
    } finally { setProvisioning(false); }
  };

  const testDv = async () => {
    try {
      const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const url = process.env.REACT_APP_DATAVERSE_URL;
      if (!enabled || !url) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL to test.' } }));
        return;
      }
      const me = await whoAmI();
      const msg = `WhoAmI OK: UserId=${me.UserId || 'n/a'} Org=${me.OrganizationId || 'n/a'}`;
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: msg } }));
    } catch (e: any) {
      const msg = typeof e?.message === 'string' ? e.message : 'Dataverse test failed';
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: msg } }));
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

  const quickDvReadTest = async () => {
    try {
      const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const url = process.env.REACT_APP_DATAVERSE_URL;
      if (!enabled || !url) {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL to run DV read test.' } }));
        return;
      }
      const t = await getDataverseToken();
      const [biz, br, docs] = await Promise.all([
        countRecords(DV_SETS.businessesSet, t).catch(() => -1),
        countRecords(DV_SETS.batchRecipientsSet, t).catch(() => -1),
        countRecords(DV_SETS.documentsSet, t).catch(() => -1)
      ]);
      const msg = `DV read ok: Businesses=${biz >= 0 ? biz : 'n/a'}, BatchRecipients=${br >= 0 ? br : 'n/a'}, Documents=${docs >= 0 ? docs : 'n/a'}`;
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: msg } }));
      try { console.info('[DV Read Test]', { businesses: biz, batchRecipients: br, documents: docs }); } catch {}
    } catch (e: any) {
      const msg = typeof e?.message === 'string' ? e.message : 'DV read test failed';
      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: msg } }));
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
        {canEdit && <button className="btn ghost" onClick={provision} title="Creates tables and columns in Dataverse" disabled={provisioning}>
          {provisioning ? 'Provisioning…' : 'Provision Dataverse Schema'}
        </button>}
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
  {canEdit && <button className="btn ghost" onClick={testDv} title="Quick WhoAmI test for Dataverse token and org">Test Dataverse</button>}
  {canEdit && <button className="btn ghost" onClick={quickDvReadTest} title="Count records from key Dataverse tables to verify live reads">Quick DV Read Test</button>}
        {canEdit && <button className="btn ghost" onClick={seedSamples} title="Insert sample rows into Batch, Document, Batch Recipient, User Progress, and Acknowledgement">Seed Sample Data</button>}
  {canEdit && <button className="btn ghost" onClick={grantCorePermissions} title="Request common Microsoft Graph permissions in one go">Grant Core Permissions</button>}
  {canEdit && <button className="btn ghost" onClick={writeTest} title="Create, fetch, and delete a test batch row to verify writes">DV Write Test</button>}
      </div>

      {/* Inline provisioning results for immediate feedback */}
      {provisionLogs && (
        <div className="card" style={{ marginTop: 8, padding: 12 }}>
          <div style={{ fontWeight: 700, color: 'var(--primary)', marginBottom: 6 }}>Provisioning Results</div>
          <div className="small muted" style={{ marginBottom: 8 }}>Review the steps below. Open the browser console for a table view.</div>
          <div style={{ display: 'grid', gap: 6 }}>
            {provisionLogs.map((log, idx) => (
              <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, border: '1px solid #eee', borderRadius: 6, padding: 8 }}>
                <span style={{ width: 10, height: 10, borderRadius: '50%', background: log.ok ? '#28a745' : '#dc3545' }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600 }}>{log.step}</div>
                  {log.detail && <div className="small muted" style={{ wordBreak: 'break-word' }}>{log.detail}</div>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
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
  const [activeTab, setActiveTab] = useState<'overview' | 'settings' | 'batch' | 'analytics' | 'dv'>('overview');
  const [healthOpen, setHealthOpen] = useState(false);
  const [healthSteps, setHealthSteps] = useState<Step[] | null>(null);
  const [granting, setGranting] = useState(false);
  const [permStatus, setPermStatus] = useState<Record<string, boolean>>({});
  const [dvStatus, setDvStatus] = useState<{ enabled: boolean; url?: string | null; userId?: string; orgId?: string; error?: string | null; lastChecked?: string | null }>({ enabled: false });
  const [accessRunning, setAccessRunning] = useState(false);
  const [accessResults, setAccessResults] = useState<Array<{ set: string; ok: boolean; status?: number; count?: number; error?: string }>>([]);
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
  useEffect(() => {
    const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
    const url = process.env.REACT_APP_DATAVERSE_URL || null;
    setDvStatus(prev => ({ ...prev, enabled, url }));
    if (!adminLight && enabled && url) {
      (async () => {
        try {
          const me = await whoAmI();
          setDvStatus({ enabled, url, userId: me.UserId, orgId: me.OrganizationId, error: null, lastChecked: new Date().toISOString() });
        } catch (e: any) {
          setDvStatus({ enabled, url, userId: undefined, orgId: undefined, error: (e?.message || 'WhoAmI failed'), lastChecked: new Date().toISOString() });
        }
      })();
    }
  }, []);
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

  // Business assignment state
  const [businesses, setBusinesses] = useState<Business[]>([]);
  const [bizLoading, setBizLoading] = useState(false);
  const [bizError, setBizError] = useState<string | null>(null);
  // Default business for all recipients (fallback)
  const [defaultBusinessId, setDefaultBusinessId] = useState<string>('');
  // Per-user business mapping keyed by email lowercased
  const [recipientBusinessMap, setRecipientBusinessMap] = useState<Record<string, string | undefined>>({});
  // Per-group business mapping keyed by group id
  const [groupBusinessMap, setGroupBusinessMap] = useState<Record<string, string | undefined>>({});

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

  // Load Businesses when Dataverse is enabled
  useEffect(() => {
    if (adminLight) return; // defer heavy load for light mode
    if (runtimeMock) return; // do not load live data in mock mode
    const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
    const url = process.env.REACT_APP_DATAVERSE_URL;
    if (!enabled || !url) return;
    (async () => {
      setBizLoading(true); setBizError(null);
      try {
        const t = await getDataverseToken();
        const list = await getBusinesses(t);
        setBusinesses(Array.isArray(list) ? list : []);
      } catch (e: any) {
        const status = e?.response?.status;
        if (status === 404) {
          setBizError('Businesses table not found. Click "Provision Dataverse Schema" or import businesses.csv, then retry.');
        } else {
          setBizError(typeof e?.message === 'string' ? e.message : 'Failed to load businesses');
        }
      } finally { setBizLoading(false); }
    })();
  }, [adminLight, runtimeMock]);

  const tabs = [
    { id: 'overview', label: 'Overview', icon: '📊' },
    { id: 'settings', label: 'Settings', icon: '⚙️' },
    { id: 'batch', label: 'Create Batch', icon: '📝' },
    { id: 'analytics', label: 'Analytics', icon: '📈' },
    { id: 'dv', label: 'Dataverse', icon: '🧩' }
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
      // 1) Persist the batch to Dataverse (best effort, only if enabled and configured)
      const dataverseEnabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
      const spEnabled = process.env.REACT_APP_ENABLE_SP_LISTS === 'true';
      let dvToken: string | undefined = undefined;
      let SETS = DV_SETS;
      try {
        if (dataverseEnabled) {
          dvToken = await getDataverseToken();
          // Resolve actual entity set names dynamically only when explicitly enabled
          if ((process.env.REACT_APP_DV_DETECT_SETS || '').toLowerCase() === 'true') {
            SETS = await resolveEntitySets();
          }
        }
      } catch {}
      // Shape example: adjust logical names to your Dataverse schema
  let createdBatchId: string | undefined;
  let spBatchId: number | null | undefined;
      if (dataverseEnabled && dvToken && process.env.REACT_APP_DATAVERSE_URL) {
        try {
          const body: any = {
            toba_name: batchForm.name,
            toba_startdate: batchForm.startDate || null,
            toba_duedate: batchForm.dueDate || null,
            toba_description: batchForm.description || null,
            toba_status: 1
          };
          const url = `${process.env.REACT_APP_DATAVERSE_URL.replace(/\/$/, '')}/api/data/v9.2/${SETS.batchesSet}`;
          let res = await fetch(url, {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${dvToken}`, 'Content-Type': 'application/json', 'Accept': 'application/json' },
            body: JSON.stringify(body)
          });
          if (!res.ok) {
            // Retry with minimal body if some custom columns don't exist in the environment
            const txt = await res.text().catch(() => '');
            console.warn('Dataverse batch creation initial failed; retrying minimal body', res.status, txt);
            const minimal: any = { toba_name: batchForm.name };
            res = await fetch(url, {
              method: 'POST', headers: { 'Authorization': `Bearer ${dvToken}`, 'Content-Type': 'application/json', 'Accept': 'application/json' },
              body: JSON.stringify(minimal)
            });
          }
          if (res.ok) {
            const loc = res.headers.get('OData-EntityId') || '';
            const match = loc.match(/[0-9a-fA-F-]{36}/);
            createdBatchId = match ? match[0] : undefined;
            logInfo('AdminPanel: created Dataverse batch', { batchId: createdBatchId });
          } else {
            console.warn('Failed to create batch in Dataverse', res.status, await res.text().catch(() => ''));
          }
        } catch (e) {
          console.warn('Dataverse batch creation error', e);
        }
      }
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

      // 2) Persist documents in Dataverse with SharePoint references (adaptive to environment schema)
  if (dataverseEnabled && dvToken && process.env.REACT_APP_DATAVERSE_URL && createdBatchId) {
        const orgBase = process.env.REACT_APP_DATAVERSE_URL.replace(/\/$/, '') + '/api/data/v9.2';
        // Helpers: detect entity logical name and attributes for the documents set
        const detectEntityAndAttrs = async (entitySet: string): Promise<{ logical?: string; attrs: Set<string> }> => {
          try {
            const defRes = await fetch(`${orgBase}/EntityDefinitions?$select=EntitySetName,LogicalName&$filter=EntitySetName eq '${entitySet}'`, { headers: { Authorization: `Bearer ${dvToken}`, Accept: 'application/json' } });
            let logical: string | undefined;
            if (defRes.ok) {
              const j = await defRes.json().catch(() => ({ value: [] }));
              logical = j?.value?.[0]?.LogicalName;
            }
            if (!logical) return { logical, attrs: new Set<string>() };
            const attrRes = await fetch(`${orgBase}/EntityDefinitions(LogicalName='${logical}')/Attributes?$select=LogicalName&$top=500`, { headers: { Authorization: `Bearer ${dvToken}`, Accept: 'application/json' } });
            if (!attrRes.ok) return { logical, attrs: new Set<string>() };
            const aj = await attrRes.json().catch(() => ({ value: [] }));
            const names: string[] = Array.isArray(aj?.value) ? aj.value.map((a: any) => a?.LogicalName).filter(Boolean) : [];
            return { logical, attrs: new Set<string>(names) };
          } catch { return { logical: undefined, attrs: new Set<string>() }; }
        };
        const pickAttr = (attrs: Set<string>, preferred?: string, candidates: string[] = [], contains: string[] = []) => {
          if (preferred && attrs.has(preferred)) return preferred;
          for (const c of candidates) if (attrs.has(c)) return c;
          const lowered = Array.from(attrs);
          for (const k of contains) {
            const hit = lowered.find(n => typeof n === 'string' && n.toLowerCase().includes(k.toLowerCase()));
            if (hit) return hit;
          }
          return '';
        };
        const postAdaptive = async (url: string, body: any) => {
          // Retry by removing invalid properties if 400 mentions them
          const maxTries = 5;
          let attempt = 0;
          let payload = { ...body };
          while (attempt < maxTries) {
            const res = await fetch(url, { method: 'POST', headers: { 'Authorization': `Bearer ${dvToken}`, 'Content-Type': 'application/json', 'Accept': 'application/json' }, body: JSON.stringify(payload) });
            if (res.ok) return true;
            const status = res.status;
            const txt = await res.text().catch(() => '');
            if (status !== 400) { console.warn('DV create failed', status, txt); return false; }
            const m = txt.match(/Invalid property '([^']+)'/);
            if (m && m[1] && payload.hasOwnProperty(m[1])) {
              delete (payload as any)[m[1]];
              attempt++;
              continue;
            }
            // If no identifiable invalid property, stop
            console.warn('DV create 400 without identifiable property', txt);
            return false;
          }
          return false;
        };

        // Detect attributes for documents and recipients once
        const docMeta = await detectEntityAndAttrs(SETS.documentsSet);
        const recMeta = await detectEntityAndAttrs(SETS.batchRecipientsSet);

        for (const d of batchForm.selectedDocuments) {
          try {
            const attrs = docMeta.attrs;
            const titleField = pickAttr(attrs, DV_ATTRS.docTitleField, ['toba_title', 'toba_name', 'name'], ['title', 'name']);
            const urlField = pickAttr(attrs, DV_ATTRS.docUrlField, ['toba_fileurl', 'toba_url'], ['url']);
            const versionField = pickAttr(attrs, DV_ATTRS.docVersionField, ['toba_version', 'version'], ['version']);
            const requiresSigField = pickAttr(attrs, DV_ATTRS.docRequiresSigField, ['toba_requiressignature'], ['sign']);

            const base: any = {};
            if (titleField) base[titleField] = d.name;
            if (urlField) base[urlField] = d.webUrl;
            if (versionField) base[versionField] = 1;
            if (requiresSigField) base[requiresSigField] = false;
            // Lookup to batch (always via configured logical name)
            base[`${DV_ATTRS.documentBatchLookup}@odata.bind`] = `/${SETS.batchesSet}(${createdBatchId})`;

            const url = `${orgBase}/${SETS.documentsSet}`;
            const ok = await postAdaptive(url, base);
            if (!ok) {
              console.warn('Failed to create document in Dataverse after adaptation', d.name);
            }
          } catch (e) {
            console.warn('Dataverse document creation error', d.name, e);
          }
        }
      }
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

      // 3) Persist Batch Recipients with Business lookups (adaptive to environment schema)
  if (dataverseEnabled && dvToken && process.env.REACT_APP_DATAVERSE_URL && createdBatchId) {
        const dvBase = process.env.REACT_APP_DATAVERSE_URL.replace(/\/$/, '') + '/api/data/v9.2';
        // Reuse helpers from documents step
        const detectEntityAndAttrs = async (entitySet: string): Promise<{ logical?: string; attrs: Set<string> }> => {
          try {
            const defRes = await fetch(`${dvBase}/EntityDefinitions?$select=EntitySetName,LogicalName&$filter=EntitySetName eq '${entitySet}'`, { headers: { Authorization: `Bearer ${dvToken}`, Accept: 'application/json' } });
            let logical: string | undefined;
            if (defRes.ok) {
              const j = await defRes.json().catch(() => ({ value: [] }));
              logical = j?.value?.[0]?.LogicalName;
            }
            if (!logical) return { logical, attrs: new Set<string>() };
            const attrRes = await fetch(`${dvBase}/EntityDefinitions(LogicalName='${logical}')/Attributes?$select=LogicalName&$top=500`, { headers: { Authorization: `Bearer ${dvToken}`, Accept: 'application/json' } });
            if (!attrRes.ok) return { logical, attrs: new Set<string>() };
            const aj = await attrRes.json().catch(() => ({ value: [] }));
            const names: string[] = Array.isArray(aj?.value) ? aj.value.map((a: any) => a?.LogicalName).filter(Boolean) : [];
            return { logical, attrs: new Set<string>(names) };
          } catch { return { logical: undefined, attrs: new Set<string>() }; }
        };
        const pickAttr = (attrs: Set<string>, preferred?: string, candidates: string[] = [], contains: string[] = []) => {
          if (preferred && attrs.has(preferred)) return preferred;
          for (const c of candidates) if (attrs.has(c)) return c;
          const lowered = Array.from(attrs);
          for (const k of contains) {
            const hit = lowered.find(n => typeof n === 'string' && n.toLowerCase().includes(k.toLowerCase()));
            if (hit) return hit;
          }
          return '';
        };
        const postAdaptive = async (url: string, body: any) => {
          const maxTries = 5;
          let attempt = 0;
          let payload = { ...body };
          while (attempt < maxTries) {
            const res = await fetch(url, { method: 'POST', headers: { 'Authorization': `Bearer ${dvToken}`, 'Content-Type': 'application/json', 'Accept': 'application/json' }, body: JSON.stringify(payload) });
            if (res.ok) return true;
            const status = res.status;
            const txt = await res.text().catch(() => '');
            if (status !== 400) { console.warn('DV create recipient failed', status, txt); return false; }
            const m = txt.match(/Invalid property '([^']+)'/);
            if (m && m[1] && payload.hasOwnProperty(m[1])) {
              delete (payload as any)[m[1]];
              attempt++;
              continue;
            }
            return false;
          }
          return false;
        };

        const recMeta = await detectEntityAndAttrs(SETS.batchRecipientsSet);
        const chooseBusiness = (emailLower: string): string | undefined => {
          // Per-user mapping first
          const direct = recipientBusinessMap[emailLower];
          if (direct) return direct;
          // Any group-origin mapping
          const origins = recipientOrigins.get(emailLower);
          if (origins) {
            for (const gid of origins) {
              const gMap = groupBusinessMap[gid];
              if (gMap) return gMap;
            }
          }
          // Fallback default
          return defaultBusinessId || undefined;
        };
        for (const r of recipients) {
          try {
            const emailLower = (r.address || '').toLowerCase();
            const u = userByEmailLower.get(emailLower);
            // Determine primary group name if any
            let primaryGroupName: string | undefined = undefined;
            const origins = recipientOrigins.get(emailLower);
            if (origins && origins.size > 0) {
              const firstGid = origins.values().next().value as string;
              const g = batchForm.selectedGroups.find(x => x.id === firstGid);
              if (g?.displayName) primaryGroupName = g.displayName;
            }
            const bId = chooseBusiness(emailLower);
            const attrs = recMeta.attrs;
            const emailField = ((): string => {
              // Prefer configured; else find a likely email attribute
              if (DV_ATTRS.batchRecipientEmailField && attrs.has(DV_ATTRS.batchRecipientEmailField)) return DV_ATTRS.batchRecipientEmailField;
              const guess = Array.from(attrs).find(a => /email|mail|upn/i.test(a));
              return guess || '';
            })();
            const userField = ((): string => {
              if (DV_ATTRS.batchRecipientUserField && attrs.has(DV_ATTRS.batchRecipientUserField)) return DV_ATTRS.batchRecipientUserField;
              const guess = Array.from(attrs).find(a => /user|upn|principal/i.test(a));
              return guess || '';
            })();
            const displayField = attrs.has(DV_ATTRS.batchRecipientDisplayNameField) ? DV_ATTRS.batchRecipientDisplayNameField : (Array.from(attrs).find(a => /displayname|name/i.test(a)) || '');
            const deptField = attrs.has(DV_ATTRS.batchRecipientDepartmentField) ? DV_ATTRS.batchRecipientDepartmentField : (Array.from(attrs).find(a => /department/i.test(a)) || '');
            const jobField = attrs.has(DV_ATTRS.batchRecipientJobTitleField) ? DV_ATTRS.batchRecipientJobTitleField : (Array.from(attrs).find(a => /jobtitle|title/i.test(a)) || '');
            const locField = attrs.has(DV_ATTRS.batchRecipientLocationField) ? DV_ATTRS.batchRecipientLocationField : (Array.from(attrs).find(a => /location|office/i.test(a)) || '');
            const pgField = attrs.has(DV_ATTRS.batchRecipientPrimaryGroupField) ? DV_ATTRS.batchRecipientPrimaryGroupField : (Array.from(attrs).find(a => /primarygroup|group/i.test(a)) || '');

            const body: any = {
              toba_name: `Recipient - ${(r.name || r.address) || ''} - ${batchForm.name}`,
              [`${DV_ATTRS.batchRecipientBatchLookup}@odata.bind`]: `/${SETS.batchesSet}(${createdBatchId})`
            };
            if (emailField) body[emailField] = emailLower;
            if (userField) body[userField] = emailLower;
            if (displayField) body[displayField] = r.name || null;
            if (deptField) body[deptField] = u?.department || null;
            if (jobField) body[jobField] = u?.jobTitle || null;
            if (locField) body[locField] = u?.officeLocation || null;
            if (pgField) body[pgField] = primaryGroupName || null;
            if (bId) body[`${DV_ATTRS.batchRecipientBusinessLookup}@odata.bind`] = `/${SETS.businessesSet}(${bId})`;

            const ok = await postAdaptive(`${dvBase}/${SETS.batchRecipientsSet}`, body);
            if (!ok) {
              console.warn('Failed to create batch recipient in Dataverse after adaptation', r.address);
            }
          } catch (e) {
            console.warn('Dataverse batch recipient creation error', r.address, e);
          }
        }
      }
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
      if (createdBatchId || spBatchId) {
        const targets = [createdBatchId ? 'Dataverse' : null, spBatchId ? 'SharePoint' : null].filter(Boolean).join(' & ');
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Batch "${batchForm.name}" created in ${targets}${batchForm.notifyByEmail ? ' and email sent' : ''}.` } }));
      } else {
        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Batch created (notifications sent) but saving to Dataverse/SharePoint failed. Check console for details.` } }));
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
      // Clear business selections
      setRecipientBusinessMap({});
      setGroupBusinessMap({});
      setDefaultBusinessId('');
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

            {/* Dataverse Status */}
            <div className="card" style={{ marginTop: 16, padding: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <div>
                  <div style={{ fontWeight: 700, color: 'var(--primary)' }}>Dataverse Status</div>
                  <div className="muted small">Environment and connectivity</div>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button className="btn ghost sm" onClick={async () => {
                    const enabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
                    const url = process.env.REACT_APP_DATAVERSE_URL || null;
                    try {
                      if (!enabled || !url) throw new Error('Dataverse disabled or URL not set');
                      const me = await whoAmI();
                      setDvStatus({ enabled, url, userId: me.UserId, orgId: me.OrganizationId, error: null, lastChecked: new Date().toISOString() });
                    } catch (e: any) {
                      setDvStatus({ enabled, url, userId: undefined, orgId: undefined, error: (e?.message || 'WhoAmI failed'), lastChecked: new Date().toISOString() });
                    }
                  }}>Re-check</button>
                  <button className="btn ghost sm" onClick={async () => {
                    try { await getDataverseToken(); window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Dataverse token acquired (if consented)' } })); }
                    catch (e) { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Dataverse token/consent failed' } })); }
                  }}>Grant DV Consent</button>
                </div>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 12 }}>
                <div style={{ border: '1px solid #eee', borderRadius: 6, padding: 10 }}>
                  <div className="small muted">Enabled</div>
                  <div style={{ fontWeight: 600 }}>{dvStatus.enabled ? 'Yes' : 'No'}</div>
                </div>
                <div style={{ border: '1px solid #eee', borderRadius: 6, padding: 10 }}>
                  <div className="small muted">Organization URL</div>
                  <div style={{ fontWeight: 600, wordBreak: 'break-all' }}>{dvStatus.url || '—'}</div>
                </div>
                <div style={{ border: '1px solid #eee', borderRadius: 6, padding: 10 }}>
                  <div className="small muted">WhoAmI User</div>
                  <div style={{ fontWeight: 600 }}>{dvStatus.userId || '—'}</div>
                </div>
                <div style={{ border: '1px solid #eee', borderRadius: 6, padding: 10 }}>
                  <div className="small muted">Organization Id</div>
                  <div style={{ fontWeight: 600 }}>{dvStatus.orgId || '—'}</div>
                </div>
                <div style={{ border: '1px solid #eee', borderRadius: 6, padding: 10 }}>
                  <div className="small muted">Last Checked</div>
                  <div style={{ fontWeight: 600 }}>{dvStatus.lastChecked ? new Date(dvStatus.lastChecked).toLocaleString() : '—'}</div>
                </div>
              </div>

              {dvStatus.error && (
                <div className="small" style={{ marginTop: 8, color: '#b71c1c' }}>Error: {dvStatus.error}</div>
              )}

              <div className="small muted" style={{ marginTop: 12 }}>
                Entity sets in use: <code>{DV_SETS.batchesSet}</code>, <code>{DV_SETS.documentsSet}</code>, <code>{DV_SETS.userAcksSet}</code>, <code>{DV_SETS.userProgressesSet}</code>
              </div>
            </div>

            {/* Dataverse Access Checker */}
            <div className="card" style={{ marginTop: 16, padding: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
                <div>
                  <div style={{ fontWeight: 700, color: 'var(--primary)' }}>Dataverse Access Checker</div>
                  <div className="muted small">Probe read privileges on key tables (no writes)</div>
                </div>
                <div style={{ display: 'flex', gap: 8 }}>
                  <button className="btn sm" disabled={accessRunning} onClick={async () => {
                    try {
                      setAccessRunning(true);
                      setAccessResults([]);
                      if (!(process.env.REACT_APP_ENABLE_DATAVERSE === 'true') || !process.env.REACT_APP_DATAVERSE_URL) {
                        window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Enable Dataverse and set REACT_APP_DATAVERSE_URL first' } }));
                        return;
                      }
                      const t = await getDataverseToken();
                      const sets = [
                        DV_SETS.businessesSet,
                        DV_SETS.batchesSet,
                        DV_SETS.documentsSet,
                        DV_SETS.batchRecipientsSet,
                        DV_SETS.userAcksSet,
                        DV_SETS.userProgressesSet
                      ];
                      const results = await Promise.all(
                        sets.map(s => probeReadAccess(s, t).catch((e: any) => ({ set: s, ok: false, status: e?.response?.status, error: String(e?.message || e) })))
                      );
                      setAccessResults(results);
                    } catch (e) {
                      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Access check failed' } }));
                    } finally {
                      setAccessRunning(false);
                    }
                  }}>{accessRunning ? 'Checking…' : 'Run Access Checks'}</button>
                  <button className="btn ghost sm" disabled={accessResults.length === 0} onClick={async () => {
                    try {
                      await navigator.clipboard.writeText(JSON.stringify(accessResults, null, 2));
                      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Access results copied to clipboard' } }));
                    } catch {
                      window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Copy failed' } }));
                    }
                  }}>Copy Results</button>
                </div>
              </div>

              {accessResults.length > 0 ? (
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 2fr', gap: 8 }}>
                  <div className="small muted" style={{ fontWeight: 600 }}>Table (Entity Set)</div>
                  <div className="small muted" style={{ fontWeight: 600 }}>Status</div>
                  <div className="small muted" style={{ fontWeight: 600 }}>Detail</div>
                  {accessResults.map((r, idx) => (
                    <React.Fragment key={idx}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <code style={{ background: '#f6f8fa', padding: '2px 6px', borderRadius: 4 }}>{r.set}</code>
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span style={{ width: 10, height: 10, borderRadius: '50%', background: r.ok ? '#28a745' : '#dc3545' }} />
                        <span className="small">{r.ok ? `OK (${r.count ?? 0})` : (r.status ? `${r.status}` : 'Error')}</span>
                      </div>
                      <div className="small muted" style={{ whiteSpace: 'pre-wrap' }}>
                        {r.ok ? 'Readable' : (
                          r.status === 401 ? 'Unauthorized (token/consent)' :
                          r.status === 403 ? 'Forbidden (missing role privilege)' :
                          r.status === 404 ? 'Not Found (entity set name mismatch)' :
                          r.error || 'Failed'
                        )}
                      </div>
                    </React.Fragment>
                  ))}
                </div>
              ) : (
                <div className="small muted">No results yet.</div>
              )}
            </div>
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

            {/* Business Assignment */}
            {(process.env.REACT_APP_ENABLE_DATAVERSE === 'true') && (
              <div className="card" style={{ padding: 16, marginBottom: 16 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ fontWeight: 700 }}>Business Assignment</div>
                    <div className="small muted">Assign each recipient to a Business. Unassigned recipients will use the default.</div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    {bizLoading && <div className="small muted">Loading businesses…</div>}
                    {process.env.REACT_APP_ENABLE_DATAVERSE === 'true' && (
                      <button className="btn ghost sm" onClick={async () => {
                        setBizLoading(true); setBizError(null);
                        try {
                          const t = await getDataverseToken();
                          const list = await getBusinesses(t);
                          setBusinesses(Array.isArray(list) ? list : []);
                        } catch (e: any) {
                          const status = e?.response?.status;
                          if (status === 404) setBizError('Businesses table not found. Provision or import businesses.csv, then retry.');
                          else setBizError(typeof e?.message === 'string' ? e.message : 'Failed to load businesses');
                        } finally { setBizLoading(false); }
                      }}>Refresh</button>
                    )}
                  </div>
                </div>
                {bizError && <div className="small" style={{ color: '#b71c1c', marginTop: 8 }}>Error: {bizError}</div>}
                {(!businesses || businesses.length === 0) && !bizLoading && (
                  <div className="small muted" style={{ marginTop: 8 }}>No businesses found. You can import them via businesses.csv or create them in Dataverse.</div>
                )}

                {businesses && businesses.length > 0 && (
                  <div style={{ marginTop: 12, display: 'grid', gap: 12 }}>
                    {/* Default business */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                      <div className="small" style={{ width: 180 }}>Default Business:</div>
                      <select value={defaultBusinessId} onChange={e => setDefaultBusinessId(e.target.value)} style={{ minWidth: 240 }}>
                        <option value="">— None —</option>
                        {businesses.map(b => <option key={b.toba_businessid} value={b.toba_businessid}>{b.toba_name}{b.toba_code ? ` (${b.toba_code})` : ''}</option>)}
                      </select>
                      <button className="btn ghost sm" onClick={() => setDefaultBusinessId('')}>Clear</button>
                    </div>

                    {/* Per-user assignments */}
                    {batchForm.selectedUsers.length > 0 && (
                      <div>
                        <div style={{ fontWeight: 600, marginBottom: 6 }}>Users</div>
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 260px', gap: 8 }}>
                          {batchForm.selectedUsers.map(u => {
                            const email = (u.mail || u.userPrincipalName || '').trim().toLowerCase();
                            return (
                              <React.Fragment key={u.id}>
                                <div>
                                  <div style={{ fontWeight: 500 }}>{u.displayName}</div>
                                  <div className="small muted">{email || u.id}</div>
                                </div>
                                <select
                                  value={recipientBusinessMap[email] || ''}
                                  onChange={e => setRecipientBusinessMap(prev => ({ ...prev, [email]: e.target.value || undefined }))}
                                >
                                  <option value="">— Default —</option>
                                  {businesses.map(b => <option key={b.toba_businessid} value={b.toba_businessid}>{b.toba_name}</option>)}
                                </select>
                              </React.Fragment>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    {/* Per-group assignments */}
                    {batchForm.selectedGroups.length > 0 && (
                      <div>
                        <div style={{ fontWeight: 600, marginBottom: 6 }}>Groups</div>
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 260px', gap: 8 }}>
                          {batchForm.selectedGroups.map(g => (
                            <React.Fragment key={g.id}>
                              <div>
                                <div style={{ fontWeight: 500 }}>{g.displayName}</div>
                                <div className="small muted">{g.memberCount || 0} members</div>
                              </div>
                              <select
                                value={groupBusinessMap[g.id] || ''}
                                onChange={e => setGroupBusinessMap(prev => ({ ...prev, [g.id]: e.target.value || undefined }))}
                              >
                                <option value="">— Default —</option>
                                {businesses.map(b => <option key={b.toba_businessid} value={b.toba_businessid}>{b.toba_name}</option>)}
                              </select>
                            </React.Fragment>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
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
        {activeTab === 'dv' && (
          <div>
            <DataverseExplorer />
          </div>
        )}
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
