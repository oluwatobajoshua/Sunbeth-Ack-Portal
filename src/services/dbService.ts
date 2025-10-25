/**
 * Data access facade for the app.
 *
 * This module is the single entry-point for fetching and mutating data in the UI.
 * It routes calls to different backends: SharePoint Lists or SQLite API.
 *
 * Contract:
 * - When REACT_APP_ENABLE_SQLITE === 'true': Uses SQLite API backend
 * - When REACT_APP_ENABLE_SP_LISTS === 'true': Uses SharePoint Lists backend
 * - Default fallback to SQLite if no backend is explicitly configured
 *
 * Add new data operations here so the rest of the app stays environment-agnostic.
 */
import { apiGet, apiPost, apiPut, apiDelete } from './api';

/**
 * Fetch the list of batches assigned to the user.
 * @param token Optional bearer token (not used with current backends)
 */
export const getBatches = async (_token?: string, userEmail?: string) => {
  const q = userEmail ? `?email=${encodeURIComponent(userEmail)}` : '';
  try { return await apiGet(`/api/batches${q}`); } catch { return [] as any[]; }
};

/**
 * Fetch the list of documents for a given batch.
 * @param batchId Batch identifier
 * @param token Optional bearer token (not used with current backends)
 */
export const getDocumentsByBatch = async (batchId: string, _token?: string) => {
  try { return await apiGet(`/api/batches/${encodeURIComponent(batchId)}/documents`); } catch { return [] as any[]; }
};

/**
 * Compute user acknowledgement progress for a batch.
 * @param batchId Batch identifier
 * @param token Optional bearer token (not used with current backends)
 * @param userId Optional user id (not used with current backends)
 */
export const getUserProgress = async (batchId: string, _token?: string, _userId?: string, userEmail?: string) => {
  const q = userEmail ? `?email=${encodeURIComponent(userEmail)}` : '';
  try { return await apiGet(`/api/batches/${encodeURIComponent(batchId)}/progress${q}`); } catch { return { acknowledged: 0, total: 0, percent: 0 } as any; }
};

/**
 * List documents in a batch, separated into acknowledged vs pending (ids only for acknowledged).
 */
export const getAcknowledgedDocIds = async (batchId: string, _token?: string, userEmail?: string): Promise<string[]> => {
  const q = userEmail ? `?email=${encodeURIComponent(userEmail)}` : '';
  try { const j = await apiGet(`/api/batches/${encodeURIComponent(batchId)}/acks${q}`); return Array.isArray(j?.ids) ? j.ids : []; } catch { return []; }
};

/**
 * Fetch a single document by id (SQLite API only for now). Returns document with batchId attached.
 */
export const getDocumentById = async (docId: string): Promise<any | null> => {
  try { return await apiGet(`/api/documents/${encodeURIComponent(docId)}`); } catch { return null; }
};

/**
 * Fetch the list of businesses.
 * @param token Optional bearer token (not used with current backends)
 */
export const getBusinesses = async (_token?: string) => {
  try { return await apiGet(`/api/businesses`); } catch { return [] as any[]; }
};

/**
 * Create a new business.
 * @param business Business data to create
 * @param token Optional bearer token (not used with current backends)
 */
export const createBusiness = async (business: { name: string; code?: string; isActive?: boolean; description?: string }, _token?: string) => {
  return await apiPost(`/api/businesses`, business);
};

/**
 * Update an existing business.
 * @param id Business ID to update
 * @param business Updated business data
 * @param token Optional bearer token (not used with current backends)
 */
export const updateBusiness = async (id: number, business: { name?: string; code?: string; isActive?: boolean; description?: string }, _token?: string) => {
  return await apiPut(`/api/businesses/${id}`, business);
};

/**
 * Delete a business.
 * @param id Business ID to delete
 * @param token Optional bearer token (not used with current backends)
 */
export const deleteBusiness = async (id: number, _token?: string) => {
  return await apiDelete(`/api/businesses/${id}`);
};

// --- Roles management (SQLite API only) ---
export type DbRole = { id: number; email: string; role: 'Admin'|'Manager'|'SuperAdmin'; createdAt?: string };

export const getRoles = async (): Promise<DbRole[]> => {
  try { const j = await apiGet(`/api/roles`); return Array.isArray(j) ? j : []; } catch { return []; }
};

export const createRole = async (email: string, role: 'Admin'|'Manager'): Promise<DbRole> => {
  return await apiPost(`/api/roles`, { email, role });
};

export const deleteRole = async (id: number): Promise<void> => {
  await apiDelete(`/api/roles/${id}`);
};
