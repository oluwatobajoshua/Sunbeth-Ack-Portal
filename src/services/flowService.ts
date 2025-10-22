import axios from 'axios';
import { info, warn, error as logError } from '../diagnostics/logger';
import { useRuntimeMock } from '../utils/runtimeMock';

const useMock = () => useRuntimeMock();

export const sendAcknowledgement = async (payload: any): Promise<void> => {
  if (useMock()) {
    // simulate network and succeed
    await new Promise(r => setTimeout(r, 250));
    info('Mock ack sent', payload);
    // persist to localStorage keyed by batchId -> array of docIds
    try {
      const key = 'mock_user_acks';
      const raw = localStorage.getItem(key);
      const map: Record<string, string[]> = raw ? JSON.parse(raw) : {};
      const batchId = payload.batchId || '1';
      map[batchId] = map[batchId] || [];
      if (!map[batchId].includes(payload.documentId)) map[batchId].push(payload.documentId);
      localStorage.setItem(key, JSON.stringify(map));
      // notify UI that a mock ack occurred
      try {
        window.dispatchEvent(new CustomEvent('mockAck', { detail: { batchId: batchId, documentId: payload.documentId } }));
        window.dispatchEvent(new CustomEvent('sunbeth:progressUpdated', { detail: { batchId: batchId, documentId: payload.documentId, mock: true } }));
      } catch (e) {
        // ignore
      }
    } catch (e) {
      warn('Failed to persist mock ack', e);
    }
    return;
  }

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
      window.dispatchEvent(new CustomEvent('sunbeth:progressUpdated', { detail: { batchId: payload.batchId, documentId: payload.documentId, mock: false } }));
    } catch {}
  } catch (e) {
    logError('Failed to send ack to flow', e);
    if (!sqliteEnabled) throw e;
  }
/**
 * Send a user acknowledgement event.
 *
 * - In mock mode: writes to localStorage under 'mock_user_acks' and emits a 'mockAck' event.
 * - In live mode: if SQLite API is enabled, persists to /api/ack; also POSTs to Flow if configured.
 */
};
