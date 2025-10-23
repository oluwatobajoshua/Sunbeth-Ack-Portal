/**
 * Notification Service (Microsoft Graph based)
 *
 * Provides helper functions to send notifications without Power Automate.
 * - Email (sendMail)
 * - (Optional) Teams 1:1 chat messages (scaffolded)
 *
 * Notes:
 * - Requires delegated Graph permissions and user consent.
 *   Email: Mail.Send
 *   Teams DM (optional): Chat.ReadWrite
 */
import { getGraphToken } from './authTokens';
import { getBrandLogoUrl, getBrandName, getBrandPrimaryColor } from '../utils/runtimeConfig';

type Recipient = { address: string; name?: string };
type MailOptions = { cc?: Recipient[]; bcc?: Recipient[] };

const chunk = <T,>(arr: T[], size: number): T[][] => {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
};

/**
 * Send a single email message to up to ~100 recipients at a time (chunked internally)
 * using the signed-in user as the sender (me/sendMail).
 */
export const sendEmail = async (
  recipients: Recipient[],
  subject: string,
  htmlBody: string,
  attachments?: Array<{ name: string; contentBytes: string; contentType?: string }>,
  options?: MailOptions
): Promise<void> => {
  if (!recipients || recipients.length === 0) return;
  const token = await getGraphToken(['Mail.Send']);

  const recipientChunks = chunk(recipients, 90); // keep well under practical limits
  for (const part of recipientChunks) {
    const message = {
      message: {
        subject,
        body: { contentType: 'HTML', content: htmlBody },
        toRecipients: part.map(r => ({ emailAddress: { address: r.address, name: r.name || r.address } })),
        ccRecipients: (options?.cc && options.cc.length > 0) ? options.cc.map(r => ({ emailAddress: { address: r.address, name: r.name || r.address } })) : undefined,
        bccRecipients: (options?.bcc && options.bcc.length > 0) ? options.bcc.map(r => ({ emailAddress: { address: r.address, name: r.name || r.address } })) : undefined,
        attachments: (attachments && attachments.length > 0) ? attachments.map(a => ({
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: a.name,
          contentType: a.contentType || 'application/octet-stream',
          contentBytes: a.contentBytes
        })) : undefined
      },
      saveToSentItems: true
    };
    const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(message)
    });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`sendMail failed: ${res.status} ${res.statusText} — ${text}`);
    }
  }
};

/**
 * Send an email and automatically split attachments across multiple messages when
 * the total size or attachment count risks hitting Graph/Exchange limits.
 * - Conservative caps: max 8 attachments or ~15 MB total per message
 * - Messages are labeled with “(part i of n)” when chunked
 */
export const sendEmailWithAttachmentChunks = async (
  recipients: Recipient[],
  subject: string,
  htmlBody: string,
  attachments?: Array<{ name: string; contentBytes: string; contentType?: string }>,
  options?: MailOptions
): Promise<void> => {
  const MAX_ATTACHMENTS = 8;
  const MAX_TOTAL_BYTES = 15 * 1024 * 1024; // ~15 MB per message (conservative)

  if (!attachments || attachments.length === 0) {
    await sendEmail(recipients, subject, htmlBody, undefined, options);
    return;
  }

  // Helper to estimate decoded byte size from base64
  const sizeOf = (b64: string) => Math.floor((b64.length * 3) / 4);

  const chunks: Array<typeof attachments> = [];
  let current: typeof attachments = [];
  let currentBytes = 0;
  for (const a of attachments) {
    const aBytes = sizeOf(a.contentBytes || '');
    const wouldExceed = current.length + 1 > MAX_ATTACHMENTS || (currentBytes + aBytes) > MAX_TOTAL_BYTES;
    if (wouldExceed && current.length > 0) {
      chunks.push(current);
      current = [];
      currentBytes = 0;
    }
    current.push(a);
    currentBytes += aBytes;
  }
  if (current.length > 0) chunks.push(current);

  if (chunks.length === 1) {
    await sendEmail(recipients, subject, htmlBody, chunks[0], options);
    return;
  }

  // Multiple parts
  for (let i = 0; i < chunks.length; i++) {
    const part = chunks[i]!;
    const partSubject = `${subject} (part ${i + 1} of ${chunks.length})`;
    const partBody = `${htmlBody}<div style="margin-top:12px;color:#666;font-size:12px">Attachments: part ${i + 1} of ${chunks.length}</div>`;
    await sendEmail(recipients, partSubject, partBody, part, options);
  }
};

/**
 * (Optional) Send a Teams 1:1 message to a set of users.
 * This requires Chat.ReadWrite delegated permission and will create a new chat per user if needed.
 * For now we scaffold the function with careful batching; you can wire it later if desired.
 */
export const sendTeamsDirectMessage = async (
  userIds: string[],
  text: string
): Promise<void> => {
  if (!userIds || userIds.length === 0) return;
  const token = await getGraphToken(['Chat.ReadWrite']);

  // Simple per-user chat create + message send
  for (const uid of userIds) {
    // Create or reuse a chat with this user
    const chatRes = await fetch('https://graph.microsoft.com/v1.0/chats', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        chatType: 'oneOnOne',
        members: [
          { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users('me')` },
          { '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${uid}')` }
        ]
      })
    });
    if (!chatRes.ok) continue; // best-effort
    const chat = await chatRes.json().catch(() => null);
    const chatId = chat?.id;
    if (!chatId) continue;

    await fetch(`https://graph.microsoft.com/v1.0/chats/${chatId}/messages`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ body: { content: text } })
    }).catch(() => {});
  }
};

/**
 * Helper to build a default email body for a new batch assignment
 */
export const buildBatchEmail = (opts: {
  appUrl: string;
  batchName: string;
  startDate?: string;
  dueDate?: string;
  description?: string;
}): { subject: string; bodyHtml: string } => {
  const brand = getBrandName();
  const logo = getBrandLogoUrl();
  const primary = getBrandPrimaryColor();
  const subject = `Action required: ${opts.batchName}`;
  const bodyHtml = `
    <div style="background:#f7f8fa;padding:24px">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #e9ecef;border-radius:10px;overflow:hidden">
        <tr>
          <td style="padding:16px 20px;border-bottom:1px solid #e9ecef;background:${primary};color:white">
            <div style="display:flex;align-items:center;gap:12px">
              ${logo ? `<img src="${logo}" alt="${brand}" style="height:28px;display:block;border:0"/>` : ''}
              <strong style="font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:16px;line-height:1">${brand}</strong>
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px 20px 8px 20px;font-family:Segoe UI,Tahoma,Arial,sans-serif;color:#111">
            <h2 style="margin:0 0 12px 0;font-size:18px;color:#111">New acknowledgement assigned</h2>
            <p style="margin:0 0 12px 0">You have been assigned: <strong>${opts.batchName}</strong>.</p>
            ${opts.description ? `<p style="margin:0 0 12px 0;color:#444">${opts.description}</p>` : ''}
            <table role="presentation" cellpadding="0" cellspacing="0" style="margin:12px 0;color:#333">
              ${opts.startDate ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Start:</td><td style="padding:4px 0"><strong>${new Date(opts.startDate).toLocaleDateString()}</strong></td></tr>` : ''}
              ${opts.dueDate ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Due:</td><td style="padding:4px 0"><strong>${new Date(opts.dueDate).toLocaleDateString()}</strong></td></tr>` : ''}
            </table>
            <p style="margin:16px 0 0 0">
              <a href="${opts.appUrl}" target="_blank" rel="noopener" style="display:inline-block;background:${primary};color:#fff;text-decoration:none;padding:10px 14px;border-radius:6px;font-weight:600">Open Acknowledgement Portal</a>
            </p>
            <p style="margin:16px 0 0 0;color:#666;font-size:12px">If the button doesn’t work, copy and paste this link into your browser:<br/><span style="word-break:break-all;color:#555">${opts.appUrl}</span></p>
          </td>
        </tr>
        <tr>
          <td style="padding:16px 20px;border-top:1px solid #e9ecef;background:#fafbfc;color:#666;font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:12px">
            <div>This message was sent by ${brand} via Microsoft Graph.</div>
            <div>Please do not reply to this automated notification.</div>
          </td>
        </tr>
      </table>
    </div>`;
  return { subject, bodyHtml };
};

/**
 * Build a simple batch completion email for HR/admins
 */
export const buildBatchCompletionEmail = (opts: {
  appUrl: string;
  batchName: string;
  completedOn?: string;
  totalRecipients?: number;
  totalDocuments?: number;
}): { subject: string; bodyHtml: string } => {
  const brand = getBrandName();
  const logo = getBrandLogoUrl();
  const primary = getBrandPrimaryColor();
  const subject = `Completed: ${opts.batchName}`;
  const bodyHtml = `
    <div style="background:#f7f8fa;padding:24px">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #e9ecef;border-radius:10px;overflow:hidden">
        <tr>
          <td style="padding:16px 20px;border-bottom:1px solid #e9ecef;background:${primary};color:white">
            <div style="display:flex;align-items:center;gap:12px">
              ${logo ? `<img src="${logo}" alt="${brand}" style="height:28px;display:block;border:0"/>` : ''}
              <strong style="font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:16px;line-height:1">${brand}</strong>
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px;font-family:Segoe UI,Tahoma,Arial,sans-serif;color:#111">
            <h2 style="margin:0 0 12px 0;font-size:18px;color:#111">Batch completed</h2>
            <p style="margin:0 0 12px 0">The batch <strong>${opts.batchName}</strong> has been fully acknowledged.</p>
            <ul style="margin:0 0 12px 18px;color:#333">
              ${opts.totalRecipients != null ? `<li><strong>Recipients:</strong> ${opts.totalRecipients}</li>` : ''}
              ${opts.totalDocuments != null ? `<li><strong>Documents per recipient:</strong> ${opts.totalDocuments}</li>` : ''}
              ${opts.completedOn ? `<li><strong>Completed on:</strong> ${new Date(opts.completedOn).toLocaleString()}</li>` : ''}
            </ul>
            <p style="margin:16px 0 0 0">
              <a href="${opts.appUrl}" target="_blank" rel="noopener" style="display:inline-block;background:${primary};color:#fff;text-decoration:none;padding:10px 14px;border-radius:6px;font-weight:600">Open Portal</a>
            </p>
          </td>
        </tr>
      </table>
    </div>`;
  return { subject, bodyHtml };
};

/**
 * Helper: fetch a URL (same-origin or via proxy) and return base64 content for Graph fileAttachment
 */
export const fetchAsBase64 = async (url: string): Promise<{ contentBytes: string; contentType?: string }> => {
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error(`Fetch failed: ${res.status} ${res.statusText}`);
  const ct = res.headers.get('content-type') || undefined;
  const buf = await res.arrayBuffer();
  const bytes = new Uint8Array(buf);
  let binary = '';
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  const contentBytes = btoa(binary);
  return { contentBytes, contentType: ct };
};

/** Build completion email for a single user's completion of a batch */
export const buildUserCompletionEmail = (opts: {
  appUrl: string;
  batchName: string;
  userEmail: string;
  userName?: string;
  completedOn?: string;
  totalDocuments?: number;
  department?: string;
  jobTitle?: string;
  location?: string;
  businessName?: string;
  primaryGroup?: string;
}): { subject: string; bodyHtml: string } => {
  const brand = getBrandName();
  const logo = getBrandLogoUrl();
  const primary = getBrandPrimaryColor();
  const subject = `Completed by ${opts.userName || opts.userEmail}: ${opts.batchName}`;
  const bodyHtml = `
    <div style="background:#f7f8fa;padding:24px">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:720px;margin:0 auto;background:#ffffff;border:1px solid #e9ecef;border-radius:10px;overflow:hidden">
        <tr>
          <td style="padding:16px 20px;border-bottom:1px solid #e9ecef;background:${primary};color:white">
            <div style="display:flex;align-items:center;gap:12px">
              ${logo ? `<img src="${logo}" alt="${brand}" style="height:28px;display:block;border:0"/>` : ''}
              <strong style="font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:16px;line-height:1">${brand}</strong>
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding:20px;font-family:Segoe UI,Tahoma,Arial,sans-serif;color:#111">
            <h2 style="margin:0 0 12px 0;font-size:18px;color:#111">User completed batch</h2>
            <p style="margin:0 0 12px 0"><strong>${opts.userName || opts.userEmail}</strong> (${opts.userEmail}) has acknowledged all documents in <strong>${opts.batchName}</strong>.</p>
            <table role="presentation" cellpadding="0" cellspacing="0" style="margin:8px 0;color:#333">
              ${opts.completedOn ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Completed on:</td><td style="padding:4px 0"><strong>${new Date(opts.completedOn).toLocaleString()}</strong></td></tr>` : ''}
              ${opts.totalDocuments != null ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Documents:</td><td style="padding:4px 0"><strong>${opts.totalDocuments}</strong></td></tr>` : ''}
              ${opts.department ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Department:</td><td style="padding:4px 0">${opts.department}</td></tr>` : ''}
              ${opts.jobTitle ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Job title:</td><td style="padding:4px 0">${opts.jobTitle}</td></tr>` : ''}
              ${opts.location ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Location:</td><td style="padding:4px 0">${opts.location}</td></tr>` : ''}
              ${opts.businessName ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Business:</td><td style="padding:4px 0">${opts.businessName}</td></tr>` : ''}
              ${opts.primaryGroup ? `<tr><td style="padding:4px 8px 4px 0;color:#666">Primary group:</td><td style="padding:4px 0">${opts.primaryGroup}</td></tr>` : ''}
            </table>
            <p style="margin:16px 0 0 0">
              <a href="${opts.appUrl}" target="_blank" rel="noopener" style="display:inline-block;background:${primary};color:#fff;text-decoration:none;padding:10px 14px;border-radius:6px;font-weight:600">Open Portal</a>
            </p>
          </td>
        </tr>
      </table>
    </div>`;
  return { subject, bodyHtml };
};
