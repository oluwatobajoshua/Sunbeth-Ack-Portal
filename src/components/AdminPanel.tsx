/* eslint-disable max-lines-per-function, complexity, react-hooks/exhaustive-deps, react-hooks/rules-of-hooks, @typescript-eslint/no-empty-function, @typescript-eslint/no-non-null-assertion, no-empty, max-lines, max-depth, no-useless-escape, global-require */
import React, { useEffect, useState } from 'react';
import { NotificationEmailsTab } from './admin/NotificationEmailsTab';
import { type SimpleDoc } from './admin/DocumentListEditor';
import { UserGroupSelector } from './admin/UserGroupSelector';
import { useAuth as useAuthCtx } from '../context/AuthContext';
import { useRBAC } from '../context/RBACContext';
// import { useAuth } from '../context/AuthContext';
import { GraphUser, GraphGroup, getGroupMembers } from '../services/graphUserService';
import AnalyticsDashboard from './AnalyticsDashboard';
import { exportAnalyticsExcel } from '../utils/excelExport';
import Modal from './Modal';
import { useFeatureFlags } from '../context/FeatureFlagsContext';
import { sendEmail, sendEmailWithAttachmentChunks, buildBatchEmail, fetchAsBase64 /*, sendTeamsDirectMessage*/ } from '../services/notificationService';
import { getGraphToken } from '../services/authTokens';
import { runAuthAndGraphCheck, Step } from '../diagnostics/health';
import { getBusinesses } from '../services/dbService';
// SharePoint Lists removed; SQLite-only mode
// SharePoint document browsing & upload (browser component handles interactions)
import BatchCreationDebug from './BatchCreationDebug';
import { alertSuccess, alertError, alertInfo, alertWarning, showToast } from '../utils/alerts';
import { busyPush, busyPop } from '../utils/busy';
import { isSQLiteEnabled, getApiBase, isAdminLight, useAdminModalSelectors as adminModalSelectorsDefault } from '../utils/runtimeConfig';
import RBACMatrix from './RBACMatrix';
import RolesManager from './admin/RolesManager';
import ExternalUsersManager from './ExternalUsersManager';
import BusinessesBulkUpload from './BusinessesBulkUpload';
import { downloadAllTemplatesExcel, downloadExternalUsersTemplateExcel, downloadExternalUsersTemplateCsv, downloadBusinessesTemplateExcel, downloadBusinessesTemplateCsv } from '../utils/importTemplates';
import AuditLogs from './AuditLogs';
import AdminSettings from './admin/AdminSettings';
import LocalLibraryPicker from './admin/LocalLibraryPicker';
import ManageBatches from './admin/ManageBatches';
import BusinessesManager from './admin/BusinessesManager';
import BatchEditor from './admin/BatchEditor';

// AdminSettings moved to ./admin/AdminSettings

// (UserGroupSelector extracted to ./admin/UserGroupSelector.tsx)

// (DocumentListEditor extracted to ./admin/DocumentListEditor.tsx)

// SharePointBrowser extracted to ./admin/SharePointBrowser
import SharePointBrowser from './admin/SharePointBrowser';

// LocalLibraryPicker moved to ./admin/LocalLibraryPicker

// Main Admin Panel Component
const AdminPanel: React.FC = () => {
  const { role, canSeeAdmin, canEditAdmin, isSuperAdmin, perms } = useRBAC();
  const { account } = useAuthCtx();
  const { externalSupport } = useFeatureFlags();
  const [activeTab, setActiveTab] = useState<'overview' | 'settings' | 'policies' | 'rbac' | 'manage' | 'batch' | 'analytics' | 'notificationEmails' | 'audit'>('overview');
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
  const adminLight = isAdminLight();
  const defaultModalToggle = ((): boolean => {
  const env = adminModalSelectorsDefault() ? 'true' : 'false';
    if (env === 'true') return true; if (env === 'false') return false; return true; // default ON to avoid mounting heavy selectors
  })();
  const [useModalSelectors, setUseModalSelectors] = useState<boolean>(() => {
    try { const v = localStorage.getItem(MODAL_TOGGLE_KEY); if (v === 'true') return true; if (v === 'false') return false; } catch {}
    return defaultModalToggle;
  });
  const [showDebugConsole, setShowDebugConsole] = useState(false);
  const [usersModalOpen, setUsersModalOpen] = useState(false);
  const [docsModalOpen, setDocsModalOpen] = useState(false);
  // Import progress (SharePoint -> Server Library)
  const [importBusy, setImportBusy] = useState(false);
  const [importTotal, setImportTotal] = useState(0);
  const [importDone, setImportDone] = useState(0);
  const [importRows, setImportRows] = useState<Array<{ name: string; status: 'saved' | 'deduped' | 'failed' }>>([]);
  // Consent Reports filters (optional)
  const [reportEmail, setReportEmail] = useState<string>('');
  const [reportSince, setReportSince] = useState<string>('');
  const [reportUntil, setReportUntil] = useState<string>('');

  // Merge helper to unify SharePoint + Local backups into a single logical selection
  const mergeDocuments = (prev: SimpleDoc[], incoming: SimpleDoc[]): SimpleDoc[] => {
    const keyOf = (d: SimpleDoc) => {
      if (d.driveId && d.itemId) return `sp:${d.driveId}:${d.itemId}`;
      if (d.localFileId != null) return `local:${d.localFileId}`;
      const name = (d.title || '').toLowerCase().trim();
      const url = (d.url || '').split('?')[0].toLowerCase().trim();
      return name || url || Math.random().toString(36).slice(2);
    };
    const map = new Map<string, SimpleDoc>();
    const mergeOne = (base: SimpleDoc, extra: SimpleDoc): SimpleDoc => {
      const prefersSharePoint = (x: SimpleDoc) => x.source === 'sharepoint' || /sharepoint\.com\//i.test(x.url);
      // Prefer SP URL as canonical when available
      const canonicalUrl = prefersSharePoint(extra) ? extra.url : prefersSharePoint(base) ? base.url : (base.url || extra.url);
      return {
        title: (extra.title && extra.title.length > (base.title || '').length) ? extra.title : base.title || extra.title || 'Document',
        url: canonicalUrl,
        version: extra.version ?? base.version,
        requiresSignature: (extra.requiresSignature ?? base.requiresSignature) || false,
        driveId: extra.driveId || base.driveId,
        itemId: extra.itemId || base.itemId,
        source: (prefersSharePoint(extra) || prefersSharePoint(base)) ? 'sharepoint' : (extra.source || base.source),
        localFileId: (extra.localFileId != null ? extra.localFileId : base.localFileId) ?? null,
        localUrl: extra.localUrl || base.localUrl || null
      };
    };
    const upsert = (d: SimpleDoc) => {
      const k = keyOf(d);
      if (map.has(k)) {
        map.set(k, mergeOne(map.get(k)!, d));
      } else {
        map.set(k, { ...d });
      }
    };
    prev.forEach(upsert);
    incoming.forEach(upsert);
    return Array.from(map.values());
  };

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
  showToast(`Loaded ${members.length} member(s) from selected group(s)`, 'success');
    } catch (e) {
  showToast('Failed to expand groups for mapping', 'error');
    }
  };
  // Remove a selected document locally and in DB when editing existing batches
  const removeSelectedDoc = async (idx: number) => {
    const doc = batchForm.selectedDocuments[idx];
    setBatchForm(prev => ({ ...prev, selectedDocuments: prev.selectedDocuments.filter((_, i) => i !== idx) }));
    if (!editingBatchId || !doc) return;
    const normalize = (u?: string | null) => (u || '').trim().toLowerCase().replace(/\/$/, '')
      .replace(/\?.*$/, ''); // strip query for robust matching
    const targetCanonical = normalize(doc.url);
    const targetLocal = normalize(doc.localUrl || undefined);
    const tryDeleteByIds = async (): Promise<boolean> => {
      try {
        const base = (getApiBase() as string);
        const res = await fetch(`${base}/api/batches/${encodeURIComponent(editingBatchId)}/documents`, { cache: 'no-store' });
        const rows: any[] = await res.json().catch(() => []);
        if (!Array.isArray(rows) || rows.length === 0) return false;
        const candidateIds: number[] = [];
        for (const r of rows) {
          const rid = Number(r.id || r.toba_documentid || r.toba_id || r.ID);
          if (!Number.isFinite(rid)) continue;
          const rCanon = normalize(r.url || r.webUrl || r.toba_originalurl || r.toba_fileurl);
          const rLocal = normalize(r.localUrl || r.toba_localurl);
          const matchUrl = (!!targetCanonical && rCanon === targetCanonical) || (!!targetLocal && rCanon === targetLocal);
          const matchLocal = (!!targetLocal && rLocal === targetLocal) || (!!targetCanonical && rLocal === targetCanonical);
          const matchDrive = (doc.driveId && doc.itemId && (r.driveId === doc.driveId || r.toba_driveid === doc.driveId) && (r.itemId === doc.itemId || r.toba_itemid === doc.itemId));
          const matchLocalId = (doc.localFileId != null) && ((r.localFileId ?? r.toba_localfileid) === doc.localFileId);
          if (matchUrl || matchLocal || matchDrive || matchLocalId) {
            candidateIds.push(rid);
          }
        }
        if (candidateIds.length === 0) return false;
        const del = await fetch(`${base}/api/batches/${encodeURIComponent(editingBatchId)}/documents`, {
          method: 'DELETE', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids: candidateIds })
        });
        if (!del.ok) return false;
        // Update originals set if we know canonical
        if (targetCanonical) {
          setOriginalDocUrls(prev => { const p = new Set(prev); p.delete(targetCanonical); return p; });
        }
        return true;
      } catch { return false; }
    };
    const tryDeleteByUrls = async (): Promise<boolean> => {
      try {
        const base = (getApiBase() as string);
        const urls: string[] = [];
        if (doc.url) urls.push(doc.url);
        if (doc.localUrl && !urls.includes(doc.localUrl)) urls.push(doc.localUrl);
        if (urls.length === 0) return false;
        const del = await fetch(`${base}/api/batches/${encodeURIComponent(editingBatchId)}/documents`, {
          method: 'DELETE', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ urls })
        });
        if (!del.ok) return false;
        if (targetCanonical) {
          setOriginalDocUrls(prev => { const p = new Set(prev); p.delete(targetCanonical); return p; });
        }
        return true;
      } catch { return false; }
    };
    // Prefer id-based deletion for precision, fallback to URLs
    const okById = await tryDeleteByIds();
    if (okById) { showToast('Removed document from batch', 'success'); return; }
    const okByUrl = await tryDeleteByUrls();
    if (okByUrl) { showToast('Removed document from batch', 'success'); return; }
    showToast('Removed locally, but server removal failed', 'warning');
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



  const tabs = (() => {
    const base: Array<{ id: string; label: string; icon: string }> = [
      { id: 'overview', label: 'Overview', icon: '📊' },
    ];
    // Settings only if allowed
    if (isSuperAdmin || perms?.manageSettings) {
      base.push({ id: 'settings', label: 'Settings', icon: '⚙️' });
      base.push({ id: 'policies', label: 'Policies', icon: '📜' } as any);
    }
    // Show Permission tab only if user can manage roles or permissions (Super Admin always)
    if (isSuperAdmin || perms?.manageRoles || perms?.managePermissions) {
      base.push({ id: 'rbac', label: 'Permission', icon: '🔐' } as any);
    }
    // Always show Notification Emails tab for super admin
    if (isSuperAdmin) {
      base.push({ id: 'notificationEmails', label: 'Notification Emails', icon: '✉️' });
    }
    base.push(
      { id: 'manage', label: 'Manage', icon: '🧰' } as any,
      // Create/edit batch only if allowed (Super Admin always)
      ...((isSuperAdmin || perms?.createBatch || perms?.editBatch) ? [{ id: 'batch', label: 'Create Batch', icon: '📝' } as any] : []),
      // Analytics only if allowed (Super Admin always)
      ...((isSuperAdmin || perms?.viewAnalytics) ? [{ id: 'analytics', label: 'Analytics', icon: '📈' } as any] : [])
    );
    // Audit Logs (Super Admin or viewDebugLogs permission)
    if (isSuperAdmin || perms?.viewDebugLogs) {
      base.push({ id: 'audit', label: 'Audit Logs', icon: '🛡️' } as any);
    }
    return base;
  })();
  const sqliteEnabled = isSQLiteEnabled();
  const [overviewStats, setOverviewStats] = useState<{ totalBatches: number; activeBatches: number; totalUsers: number; completionRate: number; overdueBatches: number; avgCompletionTime: number } | null>(null);
  type Business = { id: number; name: string; code?: string; isActive?: boolean };
  const [businesses, setBusinesses] = useState<Business[]>([]);
  const [businessMap, setBusinessMap] = useState<Record<string, number | null>>({}); // emailLower -> businessId
  const [defaultBusinessId, setDefaultBusinessId] = useState<number | ''>('');
  useEffect(() => {
    if (!sqliteEnabled) return;
    (async () => {
      try {
  const base = (getApiBase() as string);
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
      busyPush('Creating your batch...');
      // SQLite-only persistence via API
      if (!isSQLiteEnabled()) {
        await alertInfo('SQLite disabled', 'Enable SQLite (REACT_APP_ENABLE_SQLITE=true) and set REACT_APP_API_BASE.');
        return;
      }
      const base = (getApiBase() as string);

      // Validate form data
      if (!batchForm.name.trim()) {
        await alertWarning('Missing batch name', 'Batch name is required');
        return;
      }

      console.log('🚀 Starting comprehensive batch creation process...', {
        batchName: batchForm.name,
        selectedDocs: batchForm.selectedDocuments.length,
        selectedUsers: batchForm.selectedUsers.length,
        selectedGroups: batchForm.selectedGroups.length,
        isEditing: !!editingBatchId
      });

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
      // Note: email sending occurs later after successful persistence
      // Teams optional (requires Chat.ReadWrite)
      // if (batchForm.notifyByTeams) {
      //   const userIds = batchForm.selectedUsers.map(u => u.id);
      //   await sendTeamsDirectMessage(userIds, `New acknowledgement assigned: ${batchForm.name}`);
      // }

      // Pre-build documents and recipients payloads for persistence
      const allDocsPayload = batchForm.selectedDocuments.map(d => ({
        title: d.title,
        url: d.url,
        version: d.version ?? 1,
        requiresSignature: !!d.requiresSignature,
        driveId: (d as any).driveId || null,
        itemId: (d as any).itemId || null,
        source: (d as any).source || null,
        localFileId: (d as any).localFileId ?? null,
        localUrl: (d as any).localUrl ?? null
      }));
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

  // 1) Create or update batch in SQLite
  let handledRelations = false;
  let createdCounts: { documentsInserted?: number; recipientsInserted?: number } = {};
      let batchId: string | undefined;
      if (!editingBatchId) {
        // Enforce at least one doc and one recipient on create (UI guard + API contract)
        if (allDocsPayload.length === 0) {
          await alertWarning('No documents selected', 'Select at least one document');
          return;
        }
        if (recipientsPayloadAll.length === 0) {
          await alertWarning('No recipients selected', 'Select at least one recipient (user or group)');
          return;
        }

        // New atomic create: batch + documents + recipients in one transaction
        const createRes = await fetch(`${base}/api/batches/full`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            batch: {
              name: batchForm.name,
              startDate: batchForm.startDate || null,
              dueDate: batchForm.dueDate || null,
              description: batchForm.description || null,
              status: 1
            },
            documents: allDocsPayload,
            recipients: recipientsPayloadAll
          })
        });
        if (!createRes.ok) {
          let errMsg = 'batch_create_failed';
          try { const e = await createRes.json(); if (e?.error) errMsg = String(e.error); } catch { try { errMsg = await createRes.text(); } catch {} }
          console.error('Batch full create failed:', createRes.status, errMsg);
          throw new Error(errMsg || 'batch_create_failed');
        }
  const createJson = await createRes.json();
        const batchIdRaw = (createJson?.id ?? createJson?.batchId ?? createJson?.toba_batchid ?? createJson?.ID);
        batchId = typeof batchIdRaw === 'string' ? batchIdRaw : (Number.isFinite(Number(batchIdRaw)) ? String(batchIdRaw) : undefined);

        console.log('✅ DEBUG: Full batch creation success:', {
          createJson,
          finalBatchId: batchId
        });
        if (!batchId) throw new Error('batch_id_missing');
        handledRelations = true; // docs + recipients already created atomically
        createdCounts = {
          documentsInserted: Number(createJson?.documentsInserted) || undefined,
          recipientsInserted: Number(createJson?.recipientsInserted) || undefined
        };
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

      // 2) Add documents (only when editing; create handled atomically above)
      const docsToPost = !editingBatchId
        ? (handledRelations ? [] : allDocsPayload)
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

      // 3) Add recipients (only when editing; create handled atomically above)
      const recipientsPayload = editingBatchId
        ? recipientsPayloadAll.filter(r => !originalRecipientEmails.has((r.email || '').trim().toLowerCase()))
        : (handledRelations ? [] : recipientsPayloadAll);
      
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
            await alertWarning('Verification warning', 'Recipients not linked to batch (verification returned empty)');
          }
        } catch (e) {
          console.warn('Recipients verification failed', e);
        }
      }

      // Send notifications after successful persistence
      if (batchForm.notifyByEmail && recipientsToNotify.length > 0) {
        // Build attachments from all selected documents
        const attachments: Array<{ name: string; contentBytes: string; contentType?: string }> = [];
        try {
          const base = (getApiBase() as string);
          const docs = batchForm.selectedDocuments;
          for (const d of docs) {
            try {
              const title = d.title || 'document';
              const isSp = (d as any).source === 'sharepoint' || /sharepoint\.com\//i.test(d.url);
              let fileUrl = d.url;
              if (isSp) {
                try {
                  const token = await getGraphToken(['Files.Read.All','Sites.Read.All']);
                  const encoded = encodeURIComponent(d.url);
                  fileUrl = `${base}/api/proxy/graph?url=${encoded}&token=${encodeURIComponent(token)}&download=1`;
                } catch {}
              } else {
                fileUrl = `${base}/api/proxy?url=${encodeURIComponent(d.url)}`;
              }
              const { contentBytes, contentType } = await fetchAsBase64(fileUrl);
              attachments.push({ name: title, contentBytes, contentType });
            } catch (e) { /* skip this doc */ }
          }
        } catch (e) { /* non-blocking */ }
        try {
          if (attachments.length > 0) {
            await sendEmailWithAttachmentChunks(recipientsToNotify, subject, bodyHtml, attachments);
          } else {
            await sendEmail(recipientsToNotify, subject, bodyHtml, undefined);
          }
        }
        catch (e) { console.warn('Email send failed (non-blocking)', e); }
      }

      // Final feedback
      const actionWord = editingBatchId ? 'updated' : 'created';
      const countSuffix = (!editingBatchId && handledRelations)
        ? ` (${createdCounts.documentsInserted ?? 0} doc${(createdCounts.documentsInserted ?? 0) === 1 ? '' : 's'}, ${createdCounts.recipientsInserted ?? 0} recipient${(createdCounts.recipientsInserted ?? 0) === 1 ? '' : 's'})`
        : '';
      // Show success alert (overlay should be gone to let alert pop nicely)
      busyPop();
      await alertSuccess(`Batch ${actionWord}`,
        `\n<strong>${batchForm.name}</strong>${countSuffix ? `<div class=\"small muted\">${countSuffix.replace(/[()]/g,'')}</div>` : ''}` +
        (batchForm.notifyByEmail ? '<div class=\"small muted\">Email notification sent</div>' : ''),
        { showDenyButton: true, confirmButtonText: 'Great!', denyButtonText: 'Create another' }
      );

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
      await alertError('Batch save failed', (e as any)?.message || 'Failed to save batch or send notifications');
    }
    finally {
      // Ensure overlay is cleared even if we exited early
      busyPop();
    }
  };

  // Load an existing batch into the form for editing
  const startEditBatch = async (id: string) => {
    try {
      if (!sqliteEnabled) return;
  const base = (getApiBase() as string);
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
        source: d.source || d.toba_source || ((d.driveId || d.toba_driveid) ? 'sharepoint' : undefined),
        localFileId: d.localFileId || d.toba_localfileid || null,
        localUrl: d.localUrl || d.toba_localurl || null
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
      showToast('Loaded batch into editor', 'success');
    } catch (e) {
      console.error('Failed to load batch for editing', e);
      showToast('Failed to open batch for editing', 'error');
    }
  };

  // Clone an existing batch into the form for creating a new one
  const startCloneBatch = async (id: string) => {
    try {
      if (!sqliteEnabled) return;
  const base = (getApiBase() as string);
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

      const selectedDocuments = (docs || []).map((d: any) => ({
        title: d.title || d.name || d.toba_title || 'Document',
        url: d.url || d.webUrl || d.toba_fileurl,
        version: Number(d.version || d.toba_version || 1),
        requiresSignature: !!(d.requiresSignature ?? d.toba_requiressignature),
        driveId: d.driveId || d.toba_driveid || undefined,
        itemId: d.itemId || d.toba_itemid || undefined,
        source: d.source || d.toba_source || ((d.driveId || d.toba_driveid) ? 'sharepoint' : undefined),
        localFileId: d.localFileId || d.toba_localfileid || null,
        localUrl: d.localUrl || d.toba_localurl || null
      }));
      const selectedUsers = (recs || []).map((r: any) => ({ id: r.email || r.user || r.userPrincipalName || r.id || r.email, displayName: r.displayName || r.email, userPrincipalName: r.email, department: r.department, jobTitle: r.jobTitle } as any));
      const selectedGroups: GraphGroup[] = [];
      // Map user -> business
      const nextMap: Record<string, number | null> = {};
      for (const r of recs) {
        const emailLower = String(r.email || r.user || '').toLowerCase();
        if (emailLower) nextMap[emailLower] = r.businessId != null ? Number(r.businessId) : null;
      }

      setBatchForm({
        name: (String(b.toba_name || b.name || '') + ' (Copy)').trim(),
        startDate: '',
        dueDate: '',
        description: String(b.description || ''),
        selectedUsers: selectedUsers as any,
        selectedGroups,
        selectedDocuments,
        notifyByEmail: true,
        notifyByTeams: false
      });
      setBusinessMap(nextMap);
      setDefaultBusinessId('');
      setEditingBatchId(null); // new batch
      setOriginalRecipientEmails(new Set());
      setOriginalDocUrls(new Set());
      setActiveTab('batch');
  showToast('Prepared clone in editor', 'success');
    } catch (e) {
      console.error('Failed to clone batch', e);
  showToast('Failed to clone batch', 'error');
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
            {/* Intentionally removed loud role badge for a more professional, minimal header */}
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {sqliteEnabled && (
              <div className="small" title="SQLite API health" style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 8px', border: '1px solid #eee', borderRadius: 999 }}>
                <span style={{ width: 10, height: 10, borderRadius: '50%', background: apiHealth==='ok' ? '#28a745' : apiHealth==='down' ? '#dc3545' : '#ffc107' }} />
                <span>API: {apiHealth === 'ok' ? 'OK' : apiHealth === 'down' ? 'Down' : '—'}</span>
                <button className="btn ghost sm" onClick={pingApi} style={{ marginLeft: 6 }}>Refresh</button>
              </div>
            )}
            {sqliteEnabled && (
              (() => {
                try {
                  const base = (getApiBase() as string) || '';
                  if (!base) return null;
                  const host = new URL(base).host;
                  return (
                    <div className="small" title={`Backend API: ${base}`} style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '4px 8px', border: '1px solid #eee', borderRadius: 999 }}>
                      <span className="muted">Backend:</span>
                      <a href={`${base}/api/health`} target="_blank" rel="noreferrer" className="small" title="Open /api/health">
                        {host}
                      </a>
                      <button
                        className="btn ghost xs"
                        onClick={async () => {
                          try { await navigator.clipboard.writeText(base); showToast('API base copied', 'success'); }
                          catch { showToast('Copy failed', 'error'); }
                        }}
                        title="Copy API base URL"
                      >Copy</button>
                    </div>
                  );
                } catch { return null; }
              })()
            )}
            {(isSuperAdmin || perms?.exportAnalytics) && (
              <button className="btn ghost sm" onClick={async () => {
                try {
                  await exportAnalyticsExcel();
                } catch (e) { console.warn('Excel export failed', e); showToast('Excel export failed', 'error'); }
              }}>Export Excel</button>
            )}
            <button className="btn ghost sm" onClick={async () => {
              setHealthOpen(true);
              setHealthSteps(null);
              try { setHealthSteps(await runAuthAndGraphCheck()); } catch (e) { setHealthSteps([{ name: 'Health check', ok: false, detail: String(e) }]); }
            }}>System Health</button>
            {(isSuperAdmin || perms?.viewDebugLogs) && (
              <button 
                className="btn ghost sm" 
                onClick={() => setShowDebugConsole(true)}
                title="Open batch creation debug console"
              >
                🔍 Debug Logs
              </button>
            )}
            {sqliteEnabled && canEditAdmin && (
              <button className="btn ghost sm" onClick={async () => {
                try {
                  const base = (getApiBase() as string);
                  const email = account?.username || 'seed.user@sunbeth.com';
                  const res = await fetch(`${base}/api/seed?email=${encodeURIComponent(email)}`, { method: 'POST' });
                  if (!res.ok) throw new Error('seed_failed');
                  showToast('Seeded demo data', 'success');
                } catch {
                  showToast('Seed failed', 'error');
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
                      showToast('Permission prompts completed', 'success');
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

  {activeTab === 'settings' && <AdminSettings canEdit={!!(isSuperAdmin || perms?.manageSettings)} />}
  {activeTab === 'policies' && (
          <div style={{ display: 'grid', gap: 16 }}>
            <div className="card" style={{ padding: 16 }}>
              <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Recurring Policies</h3>
              <div className="small muted" style={{ marginBottom: 8 }}>Define annual or recurring acknowledgements at the document level; applies to all employees by default.</div>
              {/* Lazy import avoid unnecessary bundle growth */}
              <React.Suspense fallback={<div className="small muted">Loading policies…</div>}>
                {/* eslint-disable-next-line */}
                {React.createElement(require('./admin/Policies').default)}
              </React.Suspense>
            </div>
            {(isSuperAdmin || perms?.manageSettings) && (
              <div className="card" style={{ padding: 16 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <h3 style={{ margin: '0 0 4px 0', fontSize: 16 }}>Consent Reports</h3>
                    <div className="small muted">Download a court-ready PDF or a JSON export of consent receipts. Use filters to narrow the report.</div>
                  </div>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: 8, marginTop: 12 }}>
                  <div>
                    <label className="small" htmlFor="reportEmail">Filter Email</label>
                    <input id="reportEmail" value={reportEmail}
                      onChange={e => setReportEmail(e.target.value)}
                      placeholder="user@company.com" />
                  </div>
                  <div>
                    <label className="small" htmlFor="reportSince">Since (UTC)</label>
                    <input id="reportSince" type="datetime-local" value={reportSince}
                      onChange={e => setReportSince(e.target.value)} />
                  </div>
                  <div>
                    <label className="small" htmlFor="reportUntil">Until (UTC)</label>
                    <input id="reportUntil" type="datetime-local" value={reportUntil}
                      onChange={e => setReportUntil(e.target.value)} />
                  </div>
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 12 }}>
                  <button className="btn sm" onClick={() => {
                    try {
                      const base = (getApiBase() as string) || '';
                      const admin = account?.username || '';
                      const qs = new URLSearchParams({ adminEmail: admin });
                      if (reportEmail.trim()) qs.set('email', reportEmail.trim());
                      if (reportSince.trim()) qs.set('since', new Date(reportSince).toISOString());
                      if (reportUntil.trim()) qs.set('until', new Date(reportUntil).toISOString());
                      qs.set('format', 'pdf');
                      const url = `${base}/api/admin/consents/report?${qs.toString()}`;
                      window.open(url, '_blank');
                    } catch {}
                  }}>Download PDF Report</button>
                  <button className="btn ghost sm" onClick={() => {
                    try {
                      const base = (getApiBase() as string) || '';
                      const admin = account?.username || '';
                      const qs = new URLSearchParams({ adminEmail: admin });
                      if (reportEmail.trim()) qs.set('email', reportEmail.trim());
                      if (reportSince.trim()) qs.set('since', new Date(reportSince).toISOString());
                      if (reportUntil.trim()) qs.set('until', new Date(reportUntil).toISOString());
                      const url = `${base}/api/admin/consents/export?${qs.toString()}`;
                      window.open(url, '_blank');
                    } catch {}
                  }}>Download JSON</button>
                </div>
              </div>
            )}
          </div>
        )}

        {activeTab === 'manage' && (
          <div style={{ display: 'grid', gap: 16 }}>
            {/* Import Templates quick actions */}
            <div className="card" style={{ padding: 16 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <h3 style={{ margin: '0 0 4px 0', fontSize: 16 }}>Import Templates</h3>
                  <div className="small muted">Download enterprise-ready templates for bulk operations.</div>
                </div>
                <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <button className="btn sm" onClick={downloadAllTemplatesExcel} title="Download a workbook containing all import templates">All (Excel)</button>
                  <button className="btn ghost xs" onClick={downloadExternalUsersTemplateExcel}>External Users (Excel)</button>
                  <button className="btn ghost xs" onClick={downloadExternalUsersTemplateCsv}>External Users (CSV)</button>
                  <button className="btn ghost xs" onClick={downloadBusinessesTemplateExcel}>Businesses (Excel)</button>
                  <button className="btn ghost xs" onClick={downloadBusinessesTemplateCsv}>Businesses (CSV)</button>
                </div>
              </div>
            </div>
            {(externalSupport && (isSuperAdmin || perms?.manageRoles || perms?.manageRecipients)) && (
              <div className="card" style={{ padding: 16 }}>
                <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>External Users</h3>
                <div className="small muted" style={{ marginBottom: 8 }}>Invite, bulk upload, update, disable, or delete external users.</div>
                <ExternalUsersManager canEdit={canEditAdmin} />
              </div>
            )}
            <div className="card" style={{ padding: 16 }}>
              <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Batches</h3>
              <div className="small muted" style={{ marginBottom: 8 }}>View, edit, or delete batches. Deleting a batch removes its documents, recipients, and acknowledgements.</div>
              <ManageBatches canEdit={canEditAdmin} onEdit={(id) => startEditBatch(id)} onClone={(id) => startCloneBatch(id)} />
            </div>
            {(isSuperAdmin || perms?.manageBusinesses) && (
              <div className="card" style={{ padding: 16 }}>
                <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Businesses</h3>
                <div className="small muted" style={{ marginBottom: 8 }}>Create, edit, or delete businesses. Deleting a business will unassign it from any recipients mapped to it.</div>
                <div style={{ display: 'grid', gap: 12 }}>
                  <BusinessesBulkUploadSection />
                  <div className="divider" />
                <BusinessesManager canEdit={canEditAdmin} />
                </div>
              </div>
            )}
          </div>
        )}

  {activeTab === 'notificationEmails' && <NotificationEmailsTab />}
  {activeTab === 'rbac' && (
          <div style={{ display: 'grid', gap: 16 }}>
            {(isSuperAdmin || perms?.manageRoles) && (
              <div className="card" style={{ padding: 16 }}>
                <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Roles</h3>
                <div className="small muted" style={{ marginBottom: 8 }}>Manage Admin and Manager assignments without editing .env. Top-level access is configured via environment variables.</div>
                <RolesManager canEdit={canEditAdmin} isSuperAdmin={isSuperAdmin} />
              </div>
            )}
            {(isSuperAdmin || perms?.managePermissions) && (
              <div className="card" style={{ padding: 16 }}>
                <h3 style={{ margin: '0 0 8px 0', fontSize: 16 }}>Permissions Matrix</h3>
                <div className="small muted" style={{ marginBottom: 8 }}>Configure what each role can do, and set per-user overrides when needed.</div>
                <RBACMatrix />
              </div>
            )}
            {!(isSuperAdmin || perms?.manageRoles) && !(isSuperAdmin || perms?.managePermissions) && (
              <div className="card" style={{ padding: 16 }}>
                <div className="small muted">You don’t have permission to view Permission settings.</div>
              </div>
            )}
          </div>
        )}

        {activeTab === 'batch' && (
          <BatchEditor
            isSuperAdmin={isSuperAdmin}
            canUploadDocuments={!!(isSuperAdmin || perms?.uploadDocuments)}
            sqliteEnabled={sqliteEnabled}
            editingBatchId={editingBatchId}
            batchForm={batchForm}
            setBatchForm={setBatchForm}
            useModalSelectors={useModalSelectors}
            setUseModalSelectors={setUseModalSelectors}
            modalToggleKey={MODAL_TOGGLE_KEY}
            usersModalOpen={usersModalOpen}
            setUsersModalOpen={setUsersModalOpen}
            docsModalOpen={docsModalOpen}
            setDocsModalOpen={setDocsModalOpen}
            importBusy={importBusy}
            importTotal={importTotal}
            importDone={importDone}
            importRows={importRows}
            setImportBusy={setImportBusy}
            setImportTotal={setImportTotal}
            setImportDone={setImportDone}
            setImportRows={setImportRows}
            businesses={businesses}
            mappingUsers={mappingUsers}
            expandGroupsForMapping={expandGroupsForMapping}
            businessMap={businessMap}
            setUserBusiness={setUserBusiness}
            applyBusinessToAll={applyBusinessToAll}
            setBusinessMap={setBusinessMap}
            defaultBusinessId={defaultBusinessId}
            setDefaultBusinessId={setDefaultBusinessId}
            mergeDocuments={mergeDocuments}
            removeSelectedDoc={removeSelectedDoc}
            saveBatch={saveBatch}
          />
        )}

        {activeTab === 'audit' && (
          <div>
            <AuditLogs />
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
          <Modal open={docsModalOpen} onClose={() => setDocsModalOpen(false)} title="Documents" width={920}>
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 16 }}>
              <div style={{ display: 'grid', gap: 16 }}>
              {/* Import progress banner (modal) */}
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
                selectedDocuments: mergeDocuments(prev.selectedDocuments, docs)
              }))} />
              <SharePointBrowser canUpload={!!(isSuperAdmin || perms?.uploadDocuments)} onDocumentSelect={async (spDocs) => {
                try {
                  const base = (getApiBase() as string) || '';
                  const token = await getGraphToken(['Sites.Read.All','Files.Read.All']);
                  setImportBusy(true); setImportTotal(spDocs.length); setImportDone(0); setImportRows([]);
                  const imported: SimpleDoc[] = [];
                  let dedupedCount = 0, failed = 0;
                  for (const d of spDocs) {
                    const driveId = (d as any)?.parentReference?.driveId;
                    const itemId = (d as any)?.id;
                    const name = d.name;
                    if (!base || !driveId || !itemId || !token) {
                      imported.push({ title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint' });
                      setImportDone(v => v + 1);
                      setImportRows(rows => [...rows, { name, status: 'failed' }]);
                      failed++;
                      continue;
                    }
                    try {
                      const res = await fetch(`${base}/api/library/save-graph`, { method: 'POST', headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${token}` }, body: JSON.stringify({ driveId, itemId, name }) });
                      const j = await res.json().catch(() => null);
                      const localUrl = j?.url ? `${base}${j.url}` : undefined;
                      // Keep SharePoint link as canonical url; localUrl for server backup
                      const doc: SimpleDoc = { title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint', localFileId: j?.id ?? null, localUrl: localUrl || null };
                      imported.push(doc);
                      setImportRows(rows => [...rows, { name, status: j?.deduped ? 'deduped' : 'saved' }]);
                      if (j?.deduped) dedupedCount++;
                    } catch {
                      imported.push({ title: name, url: d.webUrl, version: 1, requiresSignature: false, driveId, itemId, source: 'sharepoint', localFileId: null, localUrl: null });
                      setImportRows(rows => [...rows, { name, status: 'failed' }]);
                      failed++;
                    } finally {
                      setImportDone(v => v + 1);
                    }
                  }
                  setBatchForm(prev => ({ ...prev, selectedDocuments: mergeDocuments(prev.selectedDocuments, imported) }));
                  showToast(`Imported ${imported.length - failed} • deduped ${dedupedCount}${failed ? ` • failed ${failed}` : ''}`, failed ? 'warning' : 'success');
                } catch (e) {
                  setBatchForm(prev => ({
                    ...prev,
                    selectedDocuments: mergeDocuments(prev.selectedDocuments, spDocs.map(d => ({ title: d.name, url: d.webUrl, version: 1, requiresSignature: false, driveId: (d as any)?.parentReference?.driveId, itemId: (d as any)?.id, source: 'sharepoint' })))
                  }));
                } finally {
                  setImportBusy(false);
                }
              }} />
              </div>
              {/* Right column: persistent selection panel */}
              <div className="card" style={{ padding: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ fontWeight: 700 }}>Selected Documents</div>
                    <div className="small muted">{batchForm.selectedDocuments.length} item(s)</div>
                  </div>
                </div>
                {batchForm.selectedDocuments.length === 0 ? (
                  <div className="small muted" style={{ marginTop: 6 }}>No documents selected yet.</div>
                ) : (
                  <div style={{ marginTop: 8, maxHeight: 440, overflowY: 'auto', display: 'grid', gap: 6 }}>
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
                        <button className="btn ghost sm" onClick={() => removeSelectedDoc(idx)} title="Remove from batch">✕</button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
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
                      showToast('Requested missing permissions', 'info');
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
const BusinessesBulkUploadSection: React.FC = () => {
  return (
    <div>
      <BusinessesBulkUpload />
    </div>
  );
};

export default AdminPanel;
 
