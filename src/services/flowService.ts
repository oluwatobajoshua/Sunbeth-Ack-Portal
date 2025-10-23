import { info, warn, error as logError } from '../diagnostics/logger';
import { getHrEmails, isSQLiteEnabled, getApiBase, getFlowAckUrl, getAdminEmails, getManagerEmails, getCompletionCcEmails, getCompletionBccEmails } from '../utils/runtimeConfig';
import { buildBatchCompletionEmail, buildUserCompletionEmail, sendEmailWithAttachmentChunks, fetchAsBase64 } from './notificationService';
import { getRoles } from './dbService';
import { getGraphToken } from './authTokens';

/**
 * Send a user acknowledgement event.
 *
 * Posts to SQLite API if enabled, and/or to Flow webhook if configured.
 */
export const sendAcknowledgement = async (payload: any): Promise<void> => {
  // If SQLite API is enabled, post to it for persistence (in addition to Flow if configured)
  const sqliteEnabled = isSQLiteEnabled();
  if (sqliteEnabled) {
    try {
      const api = getApiBase() as string;
      await fetch(`${api}/api/ack`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ batchId: payload.batchId, documentId: payload.documentId, email: (payload.userPrincipalName || payload.userEmail || payload.user || payload.userDisplay || '').toLowerCase() || payload.email || '' }) });
    } catch (e) {
      warn('SQLite ack post failed', e);
    }
  }

  const url = getFlowAckUrl() as string;
  if (!url && !sqliteEnabled) throw new Error('FLOW URL not configured');
  try {
    if (url) {
      info('Sending ack to flow', { url });
      const { default: axios } = await import('axios');
      await axios.post(url, payload, { headers: { 'Content-Type': 'application/json' } });
    }
    try {
      window.dispatchEvent(new CustomEvent('sunbeth:progressUpdated', { detail: { batchId: payload.batchId, documentId: payload.documentId } }));
    } catch {}
    // If SQLite is enabled, opportunistically check for completion
    if (sqliteEnabled && payload?.batchId) {
      try {
  const base = getApiBase() as string;
        const batchId = String(payload.batchId);
        // Load batch details
        const batchesRes = await fetch(`${base}/api/batches`);
        const batches = await batchesRes.json().catch(() => []);
        const batch = (Array.isArray(batches) ? batches : []).find((b: any) => String(b.toba_batchid || b.id) === batchId);
        // Load recipients and documents
        const [recRes, docsRes] = await Promise.all([
          fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/recipients`),
          fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/documents`)
        ]);
        const recipients = await recRes.json().catch(() => []);
        const documents = await docsRes.json().catch(() => []);
        const docCount = Array.isArray(documents) ? documents.length : 0;
        if (docCount === 0) return;

        // Per-user completion notification: only for the user in payload
        const email = String((payload.userPrincipalName || payload.userEmail || payload.user || payload.email || '')).toLowerCase();
        if (email) {
          const acksRes = await fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/acks?email=${encodeURIComponent(email)}`, { cache: 'no-store' });
          const j = await acksRes.json().catch(() => ({ ids: [] }));
          const acked = Array.isArray(j?.ids) ? j.ids.length : 0;
          if (acked >= docCount) {
            const flagKeyUser = `sunbeth:hrUserNotified:${batchId}:${email}`;
            try { if (localStorage.getItem(flagKeyUser) === '1') return; } catch {}
            let notificationEmails: string[] = [];
            try {
              if (isSQLiteEnabled()) {
                const api = getApiBase() as string;
                const res = await fetch(`${api}/api/notification-emails`);
                const j = await res.json();
                notificationEmails = Array.isArray(j?.emails) ? j.emails : [];
              }
            } catch {}
            const recipientsAll = notificationEmails
              .filter(Boolean)
              .map(a => ({ address: a }));
            if (recipientsAll.length > 0) {
              console.log('[UserCompletionEmail] Recipients:', recipientsAll.map(r => r.address));
              // Build attachments for all documents in the batch
              const attachments: Array<{ name: string; contentBytes: string; contentType?: string }> = [];
              try {
                const baseUrl = getApiBase() as string;
                // Acquire token once if any document is SharePoint-sourced
                let graphToken: string | null = null;
                for (const d of (Array.isArray(documents) ? documents : [])) {
                  try {
                    const title = String(d.toba_title || d.title || 'document');
                    const url = String(d.toba_fileurl || d.url || '');
                    const isSp = /sharepoint\.com\//i.test(url) || !!d.toba_driveid || !!d.toba_itemid || /sharepoint/i.test(String(d.toba_source || d.source || ''));
                    let fileUrl = url;
                    if (isSp) {
                      try {
                        if (!graphToken) graphToken = await getGraphToken(['Files.Read.All', 'Sites.Read.All']);
                        if (d.toba_driveid && d.toba_itemid) {
                          fileUrl = `${baseUrl}/api/proxy/graph?driveId=${encodeURIComponent(String(d.toba_driveid))}&itemId=${encodeURIComponent(String(d.toba_itemid))}&token=${encodeURIComponent(graphToken)}&download=1`;
                        } else {
                          fileUrl = `${baseUrl}/api/proxy/graph?url=${encodeURIComponent(url)}&token=${encodeURIComponent(graphToken)}&download=1`;
                        }
                      } catch {}
                    } else {
                      fileUrl = `${baseUrl}/api/proxy?url=${encodeURIComponent(url)}`;
                    }
                    const { contentBytes, contentType } = await fetchAsBase64(fileUrl);
                    attachments.push({ name: title, contentBytes, contentType });
                  } catch {}
                }
              } catch {}

              // Add per-user CSV row attachment
              try {
                const compsRes = await fetch(`${getApiBase() as string}/api/batches/${encodeURIComponent(batchId)}/completions`, { cache: 'no-store' });
                const comps = await compsRes.json().catch(() => []);
                const row = (Array.isArray(comps) ? comps : []).find((r: any) => String(r.email || '').toLowerCase() === email);
                const header = ['Email','Name','Department','JobTitle','Location','Business','PrimaryGroup','Documents','Status','CompletedAt'];
                const vals = row ? [
                  String(row.email || ''),
                  String(row.displayName || ''),
                  String(row.department || ''),
                  String(row.jobTitle || ''),
                  String(row.location || ''),
                  String(row.businessName || ''),
                  String(row.primaryGroup || ''),
                  String(row.total ?? docCount),
                  row.completed ? 'Completed' : 'Pending',
                  row.completionAt ? String(row.completionAt) : new Date().toISOString()
                ] : [email, payload.userDisplay || payload.displayName || '', '', '', '', '', '', String(docCount), 'Completed', new Date().toISOString()];
                const toCsv = (vals: string[]) => vals.map(v => (/[,"\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v)).join(',');
                const csv = [toCsv(header), toCsv(vals)].join('\r\n');
                const b64 = btoa(unescape(encodeURIComponent(csv)));
                attachments.push({ name: `${String(batch?.toba_name || batch?.name || 'batch')}-${email}-completion.csv`, contentBytes: b64, contentType: 'text/csv' });
              } catch {}

              // Enrich user profile data from recipients table and business lookup
              let recipientRow: any = null;
              try {
                recipientRow = (Array.isArray(recipients) ? recipients : []).find((r: any) => String(r.email || r.user || '').toLowerCase() === email);
              } catch {}
              let businessName: string | undefined = undefined;
              try {
                const baseBiz = getApiBase() as string;
                if (recipientRow?.businessId != null) {
                  const bizRes = await fetch(`${baseBiz}/api/businesses`, { cache: 'no-store' });
                  const biz = await bizRes.json().catch(() => []);
                  const match = (Array.isArray(biz) ? biz : []).find((b: any) => String(b.id ?? b.businessId ?? b.ID ?? b.toba_businessid) === String(recipientRow.businessId));
                  businessName = match ? String(match.name || match.Title || match.title || match.code || 'Business') : undefined;
                }
              } catch {}

              const { subject, bodyHtml } = buildUserCompletionEmail({
                appUrl: window.location.origin,
                batchName: String(batch?.toba_name || batch?.name || 'Batch'),
                userEmail: email,
                userName: payload.userDisplay || payload.displayName || undefined,
                completedOn: new Date().toISOString(),
                totalDocuments: docCount,
                department: recipientRow?.department || undefined,
                jobTitle: recipientRow?.jobTitle || undefined,
                location: recipientRow?.location || undefined,
                businessName: businessName,
                primaryGroup: recipientRow?.primaryGroup || undefined
              });
              const cc = [...getCompletionCcEmails(), email].filter(Boolean).map(a => ({ address: a }));
              const bcc = getCompletionBccEmails().map(a => ({ address: a }));
              try { await sendEmailWithAttachmentChunks(recipientsAll as any, subject, bodyHtml, attachments.length ? attachments : undefined, { cc, bcc }); } catch (err) { console.error('[UserCompletionEmail] sendEmailWithAttachmentChunks error:', err); }
              try { localStorage.setItem(flagKeyUser, '1'); } catch {}
            }
          }
        }
        // Batch-level completion notification: when all recipients have acknowledged all documents
        try {
          const flagKeyBatch = `sunbeth:hrBatchNotified:${batchId}`;
          try { if (localStorage.getItem(flagKeyBatch) === '1') { /* already notified */ return; } } catch {}
          const emails: string[] = (Array.isArray(recipients) ? recipients : [])
            .map((r: any) => String(r.email || r.user || r.userPrincipalName || '').toLowerCase())
            .filter((e: string) => !!e);
          const uniqueEmails = Array.from(new Set(emails));
          let allComplete = true;
          for (const em of uniqueEmails) {
            try {
              const res = await fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/acks?email=${encodeURIComponent(em)}`, { cache: 'no-store' });
              const jj = await res.json().catch(() => ({ ids: [] }));
              const cnt = Array.isArray(jj?.ids) ? jj.ids.length : 0;
              if (cnt < docCount) { allComplete = false; break; }
            } catch { allComplete = false; break; }
          }
          if (allComplete && uniqueEmails.length > 0 && docCount > 0) {
            const hrList = getHrEmails();
            const adminsEnv = getAdminEmails();
            const managersEnv = getManagerEmails();
            let adminsDb: string[] = [];
            let managersDb: string[] = [];
            try {
              if (isSQLiteEnabled()) {
                const roles = await getRoles();
                adminsDb = (Array.isArray(roles) ? roles : [])
                  .filter(r => String(r.role) === 'Admin' || String(r.role) === 'SuperAdmin')
                  .map(r => String(r.email).toLowerCase());
                managersDb = (Array.isArray(roles) ? roles : [])
                  .filter(r => String(r.role) === 'Manager')
                  .map(r => String(r.email).toLowerCase());
              }
            } catch {}
            const recipientsAll = Array.from(new Set([...hrList, ...adminsEnv, ...managersEnv, ...adminsDb, ...managersDb]))
              .filter(Boolean)
              .map(a => ({ address: a as string }));
            if (recipientsAll.length > 0) {
              // Attach all documents (same as per-user path)
              const attachments: Array<{ name: string; contentBytes: string; contentType?: string }> = [];
              try {
                const baseUrl = getApiBase() as string;
                let graphToken: string | null = null;
                for (const d of (Array.isArray(documents) ? documents : [])) {
                  try {
                    const title = String(d.toba_title || d.title || 'document');
                    const url = String(d.toba_fileurl || d.url || '');
                    const isSp = /sharepoint\.com\//i.test(url) || !!d.toba_driveid || !!d.toba_itemid || /sharepoint/i.test(String(d.toba_source || d.source || ''));
                    let fileUrl = url;
                    if (isSp) {
                      try {
                        if (!graphToken) graphToken = await getGraphToken(['Files.Read.All', 'Sites.Read.All']);
                        if (d.toba_driveid && d.toba_itemid) {
                          fileUrl = `${baseUrl}/api/proxy/graph?driveId=${encodeURIComponent(String(d.toba_driveid))}&itemId=${encodeURIComponent(String(d.toba_itemid))}&token=${encodeURIComponent(graphToken)}&download=1`;
                        } else {
                          fileUrl = `${baseUrl}/api/proxy/graph?url=${encodeURIComponent(url)}&token=${encodeURIComponent(graphToken)}&download=1`;
                        }
                      } catch {}
                    } else {
                      fileUrl = `${baseUrl}/api/proxy?url=${encodeURIComponent(url)}`;
                    }
                    const { contentBytes, contentType } = await fetchAsBase64(fileUrl);
                    attachments.push({ name: title, contentBytes, contentType });
                  } catch {}
                }
              } catch {}

              // Add CSV summary of recipients (all complete) using completions endpoint data
              try {
                const compsRes = await fetch(`${getApiBase() as string}/api/batches/${encodeURIComponent(batchId)}/completions`, { cache: 'no-store' });
                const comps = await compsRes.json().catch(() => []);
                const header = ['Email','Name','Department','JobTitle','Location','Business','PrimaryGroup','Documents','Status','CompletedAt'];
                const rows = (Array.isArray(comps) ? comps : [])
                  .filter((r: any) => !!r.completed)
                  .map((r: any) => [
                    String(r.email || ''),
                    String(r.displayName || ''),
                    String(r.department || ''),
                    String(r.jobTitle || ''),
                    String(r.location || ''),
                    String(r.businessName || ''),
                    String(r.primaryGroup || ''),
                    String(r.total ?? docCount),
                    r.completed ? 'Completed' : 'Pending',
                    r.completionAt ? String(r.completionAt) : ''
                  ]);
                const toCsv = (vals: string[]) => vals.map(v => (/[,"\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v)).join(',');
                const csv = [toCsv(header), ...rows.map(toCsv)].join('\r\n');
                const b64 = btoa(unescape(encodeURIComponent(csv)));
                attachments.push({ name: `${String(batch?.toba_name || batch?.name || 'batch')}-completion.csv`, contentBytes: b64, contentType: 'text/csv' });
              } catch {}

              const { subject, bodyHtml } = buildBatchCompletionEmail({
                appUrl: window.location.origin,
                batchName: String(batch?.toba_name || batch?.name || 'Batch'),
                completedOn: new Date().toISOString(),
                totalRecipients: uniqueEmails.length,
                totalDocuments: docCount
              });
              const cc = getCompletionCcEmails().map(a => ({ address: a }));
              const bcc = getCompletionBccEmails().map(a => ({ address: a }));
              try { await sendEmailWithAttachmentChunks(recipientsAll as any, subject, bodyHtml, attachments.length ? attachments : undefined, { cc, bcc }); } catch {}
              try { localStorage.setItem(flagKeyBatch, '1'); } catch {}
            }
          }
        } catch (e) {
          // Non-blocking
          warn('Batch completion notify check failed', e);
        }
      } catch (e) {
        // Non-blocking
        warn('HR notify check failed', e);
      }
    }
  } catch (e) {
    logError('Failed to send ack to flow', e);
    if (!sqliteEnabled) throw e;
  }
};
