/**
 * Data access facade for the app.
 *
 * This module is the single entry-point for fetching and mutating data in the UI.
 * It automatically routes calls to either the in-browser mock database or
 * the live Dataverse Web API based on the build-time flag REACT_APP_USE_MOCK.
 *
 * Contract:
 * - When REACT_APP_USE_MOCK === 'true':
 *   - All reads/writes go to localStorage via mockDatabaseService.
 * - When REACT_APP_USE_MOCK === 'false':
 *   - All reads go to Dataverse using the supplied MSAL token.
 *   - Writes (acks) are posted via flowService (see there for details).
 *
 * Add new data operations here so the rest of the app stays environment-agnostic.
 */
import { useRuntimeMock } from '../utils/runtimeMock';
import * as mockDb from './mockDatabaseService';
import * as dataverse from './dataverseService';
import { whoAmI } from './dataverseProvisioning';
import { getDataverseToken } from './authTokens';
import { getAckedDocIdsForUser } from './dataverseService';
import { spGetBatches, spGetDocumentsByBatch, spGetUserProgress, spGetAckedDocIdsForUser, spGetBusinesses, spGetBatchRecipients } from './spListsService';

const useMock = () => useRuntimeMock();
const dvEnabled = () => (process.env.REACT_APP_ENABLE_DATAVERSE === 'true') && !!process.env.REACT_APP_DATAVERSE_URL;
const spListsEnabled = () => (process.env.REACT_APP_ENABLE_SP_LISTS === 'true') && !!(process.env.REACT_APP_SP_SITE_ID);
const sqliteEnabled = () => (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!(process.env.REACT_APP_API_BASE);

/**
 * Fetch the list of batches assigned to the user.
 * @param token Optional bearer token (required in live mode)
 */
export const getBatches = async (token?: string, userEmail?: string) => {
  if (useMock()) return mockDb.getBatches();
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return [] as any[];
    return await res.json();
  }
  if (spListsEnabled()) {
    const rows = await spGetBatches(userEmail);
    // Map SP rows to DV-like Batch shape expected by UI
    return rows.map((r: any) => ({
      toba_batchid: String(r.id),
      toba_name: r.Title || r.title || 'Batch',
      toba_startdate: r.startDate || undefined,
      toba_duedate: r.dueDate || undefined,
      toba_status: (r.status != null ? String(r.status) : undefined)
    }));
  }
  if (!dvEnabled()) return mockDb.getBatches();
  try {
    const t = token || await getDataverseToken();
    return await dataverse.getBatches(t, userEmail);
  } catch (e) {
    // Soft-fail in UI: return empty to avoid error banners
    return [] as any[];
  }
};

/**
 * Fetch the list of documents for a given batch.
 * @param batchId Batch identifier (Dataverse GUID string in live mode)
 * @param token Optional bearer token (required in live mode)
 */
export const getDocumentsByBatch = async (batchId: string, token?: string) => {
  if (useMock()) return mockDb.getDocumentsByBatch(batchId);
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const res = await fetch(`${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/documents`);
    if (!res.ok) return [] as any[];
    return await res.json();
  }
  if (spListsEnabled()) {
    const rows = await spGetDocumentsByBatch(Number(batchId));
    return rows.map((r: any) => ({
      toba_documentid: String(r.id),
      toba_title: r.title || r.Title || 'Document',
      toba_version: r.version != null ? String(r.version) : undefined,
      toba_requiressignature: !!r.requiresSignature,
      toba_fileurl: r.url
    }));
  }
  if (!dvEnabled()) return mockDb.getDocumentsByBatch(batchId);
  try {
    const t = token || await getDataverseToken();
    return await dataverse.getDocumentsByBatch(batchId, t);
  } catch (e) {
    return [] as any[];
  }
};

/**
 * Compute user acknowledgement progress for a batch.
 * In mock mode this is computed from localStorage; in live mode this queries Dataverse.
 * @param batchId Batch identifier
 * @param token Optional bearer token (required in live mode)
 * @param userId Optional user id if your schema requires it
 */
export const getUserProgress = async (batchId: string, token?: string, userId?: string, userEmail?: string) => {
  if (useMock()) return mockDb.getUserProgress(batchId);
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/progress${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return { acknowledged: 0, total: 0, percent: 0 } as any;
    return await res.json();
  }
  if (spListsEnabled()) return spGetUserProgress(Number(batchId), userEmail);
  if (!dvEnabled()) return mockDb.getUserProgress(batchId);
  const t = token || await getDataverseToken();
  // If neither a Dataverse user id nor email is provided, try to resolve DV user id via WhoAmI.
  let uid = userId;
  try {
    if (!uid && !userEmail) {
      const me = await whoAmI();
      uid = me.UserId || undefined;
    }
  } catch { /* ignore, fallback to undefined */ }
  return dataverse.getUserProgress(batchId, t, uid, userEmail);
};

/**
 * Seed mock data into localStorage. No-op in live mode.
 */
export const seed = (batches?: any, docs?: any, acks?: any) => {
  if (useMock()) return mockDb.seed(batches, docs, acks);
  // no-op for live
  return;
};

/**
 * Clear mock data from localStorage. No-op in live mode.
 */
export const clear = () => {
  if (useMock()) return mockDb.clear();
  return;
};

/**
 * List documents in a batch, separated into acknowledged vs pending (ids only for acknowledged).
 * In mock mode, uses localStorage acks. In live mode, this should query acknowledgements table.
 */
export const getAcknowledgedDocIds = async (batchId: string, token?: string, userEmail?: string): Promise<string[]> => {
  if (useMock()) return (await import('./mockDatabaseService')).getAcknowledgedDocIds(batchId);
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/acks${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return [];
    const j = await res.json();
    return Array.isArray(j?.ids) ? j.ids : [];
  }
  if (spListsEnabled()) return (await spGetAckedDocIdsForUser(Number(batchId), userEmail)).map(String);
  try {
    const t = token || await getDataverseToken();
    return await getAckedDocIdsForUser(batchId, t, userEmail);
  } catch {
    return [];
  }
};
