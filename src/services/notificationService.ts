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

type Recipient = { address: string; name?: string };

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
  htmlBody: string
): Promise<void> => {
  if (!recipients || recipients.length === 0) return;
  const token = await getGraphToken(['Mail.Send']);

  const recipientChunks = chunk(recipients, 90); // keep well under practical limits
  for (const part of recipientChunks) {
    const message = {
      message: {
        subject,
        body: { contentType: 'HTML', content: htmlBody },
        toRecipients: part.map(r => ({ emailAddress: { address: r.address, name: r.name || r.address } }))
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
  const subject = `New Acknowledgement Assigned: ${opts.batchName}`;
  const bodyHtml = `
    <div style="font-family:Segoe UI,Tahoma,Arial,sans-serif;font-size:14px;color:#111">
      <p>Hello,</p>
      <p>You have been assigned a new acknowledgement batch: <strong>${opts.batchName}</strong>.</p>
      ${opts.description ? `<p>${opts.description}</p>` : ''}
      <ul>
        ${opts.startDate ? `<li><strong>Start:</strong> ${new Date(opts.startDate).toLocaleDateString()}</li>` : ''}
        ${opts.dueDate ? `<li><strong>Due:</strong> ${new Date(opts.dueDate).toLocaleDateString()}</li>` : ''}
      </ul>
      <p>
        Please open the portal to review and acknowledge the assigned documents:<br/>
        <a href="${opts.appUrl}" target="_blank" rel="noopener">Open Sunbeth Acknowledgement Portal</a>
      </p>
      <p style="color:#666">This message was sent via Microsoft Graph. Do not reply.</p>
    </div>
  `;
  return { subject, bodyHtml };
};
