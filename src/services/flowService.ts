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
  // Always post to SQLite API for persistence
  try {
    const api = getApiBase() as string;
    const ackPayload = {
      batchId: payload.batchId,
      documentId: payload.documentId,
      email: (payload.userPrincipalName || payload.userEmail || payload.user || payload.userDisplay || '').toLowerCase() || payload.email || ''
    };
    console.log('[sendAcknowledgement] Payload to /api/ack:', ackPayload);
    const ackRes = await fetch(`${api}/api/ack`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(ackPayload) });
    const ackResText = await ackRes.text().catch(() => '');
    console.log('[sendAcknowledgement] /api/ack response:', ackRes.status, ackResText);
    if (!ackRes.ok) warn('SQLite ack post failed', { status: ackRes.status, body: ackResText });
  } catch (e) {
    warn('SQLite ack post failed (exception)', e);
  }

  // Always check for completion and send admin notification directly (no FLOW dependency)
  try {
    const base = getApiBase() as string;
    const batchId = String(payload.batchId);
    console.log('[sendAcknowledgement] Checking completion for batch:', batchId);
    // Load batch details
    const batchesRes = await fetch(`${base}/api/batches`);
    const batches = await batchesRes.json().catch(() => []);
    console.log('[sendAcknowledgement] Loaded batches:', batches);
    const batch = (Array.isArray(batches) ? batches : []).find((b: any) => String(b.toba_batchid || b.id) === batchId);
    // Load recipients and documents
    const [recRes, docsRes] = await Promise.all([
      fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/recipients`),
      fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/documents`)
    ]);
    const recipients = await recRes.json().catch(() => []);
    const documents = await docsRes.json().catch(() => []);
    console.log('[sendAcknowledgement] Recipients:', recipients);
    console.log('[sendAcknowledgement] Documents:', documents);
    const docCount = Array.isArray(documents) ? documents.length : 0;
    if (docCount === 0) {
      console.warn('[sendAcknowledgement] No documents found for batch:', batchId);
      return;
    }

    // Per-user completion notification: only for the user in payload
    const email = String((payload.userPrincipalName || payload.userEmail || payload.user || payload.email || '')).toLowerCase();
    if (email) {
      const acksRes = await fetch(`${base}/api/batches/${encodeURIComponent(batchId)}/acks?email=${encodeURIComponent(email)}`, { cache: 'no-store' });
      const j = await acksRes.json().catch(() => ({ ids: [] }));
      const acked = Array.isArray(j?.ids) ? j.ids.length : 0;
      console.log('[sendAcknowledgement] User acks:', j?.ids, 'acked:', acked, 'docCount:', docCount);
      if (acked >= docCount) {
        const flagKeyUser = `sunbeth:hrUserNotified:${batchId}:${email}`;
        try { if (localStorage.getItem(flagKeyUser) === '1') { console.log('[sendAcknowledgement] Already notified for user:', flagKeyUser); return; } } catch {}
        let notificationEmails: string[] = [];
        try {
          const api = getApiBase() as string;
          const res = await fetch(`${api}/api/notification-emails`);
          const j = await res.json();
          notificationEmails = Array.isArray(j?.emails) ? j.emails : [];
          console.log('[sendAcknowledgement] Loaded notificationEmails:', notificationEmails);
        } catch (err) {
          console.warn('[sendAcknowledgement] Failed to load notificationEmails:', err);
        }
        const recipientsAll = notificationEmails
          .filter(Boolean)
          .map(a => ({ address: a }));
        if (recipientsAll.length > 0) {
          console.log('[UserCompletionEmail] Recipients:', recipientsAll.map(r => r.address));
          if (!recipientsAll.length) {
            console.warn('[UserCompletionEmail] No admin recipients found in notificationEmails:', notificationEmails);
          }
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
                  } catch (err) {
                    console.warn('[sendAcknowledgement] Failed to get Graph token or file for doc:', d, err);
                  }
                } else {
                  fileUrl = `${baseUrl}/api/proxy?url=${encodeURIComponent(url)}`;
                }
                const { contentBytes, contentType } = await fetchAsBase64(fileUrl);
                attachments.push({ name: title, contentBytes, contentType });
              } catch (err) {
                console.warn('[sendAcknowledgement] Failed to build attachment for doc:', d, err);
              }
            }
          } catch (err) {
            console.warn('[sendAcknowledgement] Failed to build attachments:', err);
          }

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
            console.log('[sendAcknowledgement] Built CSV attachment for user:', email);
          } catch (err) {
            console.warn('[sendAcknowledgement] Failed to build CSV attachment:', err);
          }

          // Enrich user profile data from recipients table and business lookup
          let recipientRow: any = null;
          try {
            recipientRow = (Array.isArray(recipients) ? recipients : []).find((r: any) => String(r.email || r.user || '').toLowerCase() === email);
            console.log('[sendAcknowledgement] Found recipientRow:', recipientRow);
          } catch (err) {
            console.warn('[sendAcknowledgement] Failed to find recipientRow:', err);
          }
          let businessName: string | undefined = undefined;
          try {
            const baseBiz = getApiBase() as string;
            if (recipientRow?.businessId != null) {
              const bizRes = await fetch(`${baseBiz}/api/businesses`, { cache: 'no-store' });
              const biz = await bizRes.json().catch(() => []);
              const match = (Array.isArray(biz) ? biz : []).find((b: any) => String(b.id ?? b.businessId ?? b.ID ?? b.toba_businessid) === String(recipientRow.businessId));
              businessName = match ? String(match.name || match.Title || match.title || match.code || 'Business') : undefined;
              console.log('[sendAcknowledgement] Found businessName:', businessName);
            }
          } catch (err) {
            console.warn('[sendAcknowledgement] Failed to find businessName:', err);
          }

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
          console.log('[sendAcknowledgement] Email subject:', subject);
          // Send email
          const cc = [...getCompletionCcEmails(), email].filter(Boolean).map(a => ({ address: a }));
          const bcc = getCompletionBccEmails().map(a => ({ address: a }));
          try {
            await sendEmailWithAttachmentChunks(recipientsAll as any, subject, bodyHtml, attachments.length ? attachments : undefined, { cc, bcc });
            console.log('[UserCompletionEmail] Email sent to admins:', recipientsAll.map(r => r.address));
          } catch (err) {
            console.error('[UserCompletionEmail] sendEmailWithAttachmentChunks error:', err);
            let errMsg = '';
            if (err && typeof err === 'object' && 'message' in err) {
              errMsg = (err as any).message;
            } else {
              try { errMsg = JSON.stringify(err); } catch { errMsg = String(err); }
            }
            alert('[UserCompletionEmail] Failed to send admin notification: ' + errMsg);
          }
          try { localStorage.setItem(flagKeyUser, '1'); } catch {}
        } else {
          console.warn('[sendAcknowledgement] No admin recipients found in notificationEmails:', notificationEmails);
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
        // ...existing code...
      }
    } catch (e) {
      // Non-blocking
      warn('Batch completion notify check failed', e);
    }
  } catch (e) {
    // Non-blocking
    warn('HR notify check failed', e);
  }
};
