import axios from 'axios';
import { info, warn, error as logError } from '../diagnostics/logger';

/**
 * Send a user acknowledgement event.
 *
 * Posts to SQLite API if enabled, and/or to Flow webhook if configured.
 */
export const sendAcknowledgement = async (payload: any): Promise<void> => {
  // If SQLite API is enabled, post to it for persistence (in addition to Flow if configured)
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  if (sqliteEnabled) {
    try {
      const api = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
      await fetch(`${api}/api/ack`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ batchId: payload.batchId, documentId: payload.documentId, email: (payload.userPrincipalName || payload.userEmail || payload.user || payload.userDisplay || '').toLowerCase() || payload.email || '' }) });
    } catch (e) {
      warn('SQLite ack post failed', e);
    }
  }

  const url = process.env.REACT_APP_FLOW_CREATE_USER_ACK_URL as string;
  if (!url && !sqliteEnabled) throw new Error('FLOW URL not configured');
  try {
    if (url) {
      info('Sending ack to flow', { url });
      await axios.post(url, payload, { headers: { 'Content-Type': 'application/json' } });
    }
    try {
      window.dispatchEvent(new CustomEvent('sunbeth:progressUpdated', { detail: { batchId: payload.batchId, documentId: payload.documentId } }));
    } catch {}
  } catch (e) {
    logError('Failed to send ack to flow', e);
    if (!sqliteEnabled) throw e;
  }
};
