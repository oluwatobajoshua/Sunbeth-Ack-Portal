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
import { spGetBatches, spGetDocumentsByBatch, spGetUserProgress, spGetAckedDocIdsForUser, spGetBusinesses, spGetBatchRecipients } from './spListsService';

const spListsEnabled = () => (process.env.REACT_APP_ENABLE_SP_LISTS === 'true') && !!(process.env.REACT_APP_SP_SITE_ID);
const sqliteEnabled = () => (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!(process.env.REACT_APP_API_BASE);

/**
 * Fetch the list of batches assigned to the user.
 * @param token Optional bearer token (not used with current backends)
 */
export const getBatches = async (token?: string, userEmail?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return [] as any[];
    return await res.json();
  }
  
  if (spListsEnabled()) {
    const rows = await spGetBatches(userEmail);
    // Map SP rows to batch shape expected by UI
    return rows.map((r: any) => ({
      toba_batchid: String(r.id),
      toba_name: r.Title || r.title || 'Batch',
      toba_startdate: r.startDate || undefined,
      toba_duedate: r.dueDate || undefined,
      toba_status: (r.status != null ? String(r.status) : undefined)
    }));
  }
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const url = `${base.replace(/\/$/, '')}/api/batches${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
  const res = await fetch(url);
  if (!res.ok) return [] as any[];
  return await res.json();
};

/**
 * Fetch the list of documents for a given batch.
 * @param batchId Batch identifier
 * @param token Optional bearer token (not used with current backends)
 */
export const getDocumentsByBatch = async (batchId: string, token?: string) => {
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
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const res = await fetch(`${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/documents`);
  if (!res.ok) return [] as any[];
  return await res.json();
};

/**
 * Compute user acknowledgement progress for a batch.
 * @param batchId Batch identifier
 * @param token Optional bearer token (not used with current backends)
 * @param userId Optional user id (not used with current backends)
 */
export const getUserProgress = async (batchId: string, token?: string, userId?: string, userEmail?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/progress${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return { acknowledged: 0, total: 0, percent: 0 } as any;
    return await res.json();
  }
  
  if (spListsEnabled()) return spGetUserProgress(Number(batchId), userEmail);
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/progress${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
  const res = await fetch(url);
  if (!res.ok) return { acknowledged: 0, total: 0, percent: 0 } as any;
  return await res.json();
};

/**
 * List documents in a batch, separated into acknowledged vs pending (ids only for acknowledged).
 */
export const getAcknowledgedDocIds = async (batchId: string, token?: string, userEmail?: string): Promise<string[]> => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/acks${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
    const res = await fetch(url);
    if (!res.ok) return [];
    const j = await res.json();
    return Array.isArray(j?.ids) ? j.ids : [];
  }
  
  if (spListsEnabled()) return (await spGetAckedDocIdsForUser(Number(batchId), userEmail)).map(String);
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const url = `${base.replace(/\/$/, '')}/api/batches/${encodeURIComponent(batchId)}/acks${userEmail ? `?email=${encodeURIComponent(userEmail)}` : ''}`;
  const res = await fetch(url);
  if (!res.ok) return [];
  const j = await res.json();
  return Array.isArray(j?.ids) ? j.ids : [];
};

/**
 * Fetch a single document by id (SQLite API only for now). Returns document with batchId attached.
 */
export const getDocumentById = async (docId: string): Promise<any | null> => {
  const base = (process.env.REACT_APP_API_BASE || 'http://localhost:4000').replace(/\/$/, '');
  try {
    const res = await fetch(`${base}/api/documents/${encodeURIComponent(docId)}`);
    if (!res.ok) return null;
    const j = await res.json();
    return j;
  } catch {
    return null;
  }
};

/**
 * Fetch the list of businesses.
 * @param token Optional bearer token (not used with current backends)
 */
export const getBusinesses = async (token?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses`);
    if (!res.ok) return [] as any[];
    return await res.json();
  }
  
  if (spListsEnabled()) {
    const rows = await spGetBusinesses();
    // Map SP rows to business shape expected by UI
    return rows.map((r: any) => ({
      id: r.id,
      name: r.Title || r.title || r.name || 'Business',
      code: r.code || null,
      isActive: r.isActive !== false,
      description: r.description || null
    }));
  }
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses`);
  if (!res.ok) return [] as any[];
  return await res.json();
};

/**
 * Create a new business.
 * @param business Business data to create
 * @param token Optional bearer token (not used with current backends)
 */
export const createBusiness = async (business: { name: string; code?: string; isActive?: boolean; description?: string }, token?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(business)
    });
    if (!res.ok) throw new Error('Failed to create business');
    return await res.json();
  }
  
  if (spListsEnabled()) {
    // Would need to implement SharePoint creation - not implemented yet
    throw new Error('Business creation not implemented for SharePoint Lists');
  }
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(business)
  });
  if (!res.ok) throw new Error('Failed to create business');
  return await res.json();
};

/**
 * Update an existing business.
 * @param id Business ID to update
 * @param business Updated business data
 * @param token Optional bearer token (not used with current backends)
 */
export const updateBusiness = async (id: number, business: { name?: string; code?: string; isActive?: boolean; description?: string }, token?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(business)
    });
    if (!res.ok) throw new Error('Failed to update business');
    return await res.json();
  }
  
  if (spListsEnabled()) {
    // Would need to implement SharePoint update - not implemented yet
    throw new Error('Business update not implemented for SharePoint Lists');
  }
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses/${id}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(business)
  });
  if (!res.ok) throw new Error('Failed to update business');
  return await res.json();
};

/**
 * Delete a business.
 * @param id Business ID to delete
 * @param token Optional bearer token (not used with current backends)
 */
export const deleteBusiness = async (id: number, token?: string) => {
  if (sqliteEnabled()) {
    const base = process.env.REACT_APP_API_BASE as string;
    const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses/${id}`, {
      method: 'DELETE'
    });
    if (!res.ok) throw new Error('Failed to delete business');
    return await res.json();
  }
  
  if (spListsEnabled()) {
    // Would need to implement SharePoint delete - not implemented yet
    throw new Error('Business deletion not implemented for SharePoint Lists');
  }
  
  // Default fallback to SQLite if not explicitly configured
  const base = process.env.REACT_APP_API_BASE || 'http://localhost:4000';
  const res = await fetch(`${base.replace(/\/$/, '')}/api/businesses/${id}`, {
    method: 'DELETE'
  });
  if (!res.ok) throw new Error('Failed to delete business');
  return await res.json();
};

// --- Roles management (SQLite API only) ---
export type DbRole = { id: number; email: string; role: 'Admin'|'Manager'|'SuperAdmin'; createdAt?: string };

export const getRoles = async (): Promise<DbRole[]> => {
  const base = (process.env.REACT_APP_API_BASE || 'http://localhost:4000').replace(/\/$/, '');
  const res = await fetch(`${base}/api/roles`);
  if (!res.ok) return [];
  const j = await res.json();
  return Array.isArray(j) ? j : [];
};

export const createRole = async (email: string, role: 'Admin'|'Manager'): Promise<DbRole> => {
  const base = (process.env.REACT_APP_API_BASE || 'http://localhost:4000').replace(/\/$/, '');
  const res = await fetch(`${base}/api/roles`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, role })
  });
  if (!res.ok) throw new Error('Failed to create role');
  return await res.json();
};

export const deleteRole = async (id: number): Promise<void> => {
  const base = (process.env.REACT_APP_API_BASE || 'http://localhost:4000').replace(/\/$/, '');
  const res = await fetch(`${base}/api/roles/${id}`, { method: 'DELETE' });
  if (!res.ok) throw new Error('Failed to delete role');
};
