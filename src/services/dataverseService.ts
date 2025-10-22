import axios from 'axios';
import { info, warn, error as logError } from '../diagnostics/logger';
import type { Batch, Doc, Business } from '../types/models';
import { DV_SETS } from './dataverseConfig';
import { DV_ATTRS } from './dataverseConfig';
import { DV_FLAGS } from './dataverseConfig';

const baseUrl = `${process.env.REACT_APP_DATAVERSE_URL?.replace(/\/$/, '')}/api/data/v9.2`;

/**
 * Dataverse Web API client (live mode only).
 *
 * This module assumes a valid Bearer token is provided by MSAL and focuses on
 * constructing the correct Dataverse OData requests.
 *
 * Endpoints used:
 * - GET {org}/api/data/v9.2/toba_batches?$select=...
 * - GET {org}/api/data/v9.2/toba_documents?$select=...&$filter=_toba_batch_value eq {batchId}
 *
 * Notes:
 * - Adjust entity logical names/attributes to match your Dataverse schema.
 * - For filters on GUIDs, you may need to wrap with quotes depending on the environment: ... eq {GUID}
 *   If needed, switch to ... eq ${batchId.replace(/[{}/]/g,'')}
 */
export const getBatches = async (token?: string, userEmail?: string): Promise<Batch[]> => {
  try {
    // If a user email is provided, fetch only batches assigned to that user via Batch Recipients
    if (userEmail) {
      const emailSafe = String(userEmail).replace(/'/g, "''");
      // 1) Get recipient rows for this user -> extract batch ids
      const recUrl = `${baseUrl}/${DV_SETS.batchRecipientsSet}?$select=_toba_batch_value,${DV_ATTRS.batchRecipientEmailField}&$filter=${DV_ATTRS.batchRecipientEmailField} eq '${emailSafe}'`;
      let rows: Array<any> = [];
      try {
        const recRes = await axios.get(recUrl, {
          headers: {
            Authorization: `Bearer ${token}`,
            Accept: 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            Prefer: 'odata.include-annotations="*"'
          }
        });
        rows = Array.isArray(recRes.data?.value) ? recRes.data.value : [];
      } catch (e: any) {
        const status = e?.response?.status;
        if ((status === 401 || status === 403) && DV_FLAGS.fallbackAllBatchesOn401) {
          warn('getBatches: recipients unauthorized, falling back to all batches', { status });
          const resAll = await axios.get(`${baseUrl}/${DV_SETS.batchesSet}?$select=toba_batchid,toba_name,toba_startdate,toba_duedate,toba_status`, {
            headers: {
              Authorization: `Bearer ${token}`,
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              Prefer: 'odata.include-annotations="*"'
            }
          });
          const all = Array.isArray(resAll.data?.value) ? resAll.data.value : [];
          info('getBatches (fallback all) fetched from dataverse', { count: all.length });
          try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Showing all batches (recipients unauthorized)' } })); } catch {}
          return all;
        }
        if (status === 401 || status === 403) {
          warn('getBatches: unauthorized/forbidden', { status });
          return [];
        }
        throw e;
      }
      const batchIds = Array.from(new Set(rows.map(r => r?._toba_batch_value).filter(Boolean)));
      if (batchIds.length === 0) return [];
      // 2) Fetch batches by id (chunk if needed)
      const select = '$select=toba_batchid,toba_name,toba_startdate,toba_duedate,toba_status';
      const fetchOne = async (id: string) => {
        const url = `${baseUrl}/${DV_SETS.batchesSet}?${select}&$filter=toba_batchid eq ${String(id).replace(/[{}]/g,'')}`;
        const r = await axios.get(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'odata.include-annotations="*"' } });
        const arr = Array.isArray(r.data?.value) ? r.data.value : [];
        return arr[0];
      };
      const results: Batch[] = [];
      // Limit concurrency a bit
      const chunkSize = 5;
      for (let i = 0; i < batchIds.length; i += chunkSize) {
        const chunk = batchIds.slice(i, i + chunkSize);
        const vals = await Promise.all(chunk.map(id => fetchOne(id).catch(() => null)));
        for (const v of vals) if (v) results.push(v);
      }
      info('getBatches (assigned) fetched from dataverse', { count: results.length });
      return results;
    }

    // Fallback: fetch all batches (admin-like scenario)
    const res = await axios.get(`${baseUrl}/${DV_SETS.batchesSet}?$select=toba_batchid,toba_name,toba_startdate,toba_duedate,toba_status`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="*"'
      }
    });
    info('getBatches fetched from dataverse', { url: `${baseUrl}/${DV_SETS.batchesSet}` });
    return res.data.value;
  } catch (e: any) {
    const status = e?.response?.status;
    if (status === 404) {
      warn('getBatches: entity set not found', { set: DV_SETS.batchesSet });
      return [];
    }
    if (status === 401 || status === 403) {
      warn('getBatches: unauthorized/forbidden', { status });
      return [];
    }
    throw e;
  }
};

export const getDocumentsByBatch = async (batchId: string, token?: string): Promise<Doc[]> => {
  // Note: OData filter adapted for Dataverse; ensure proper quoting
  const bid = (batchId || '').replace(/[{}]/g, '');
  try {
    const res = await axios.get(`${baseUrl}/${DV_SETS.documentsSet}?$select=toba_documentid,toba_title,toba_version,toba_requiressignature,toba_fileurl&$filter=_toba_batch_value eq ${bid}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          Prefer: 'odata.include-annotations="*"'
        }
      });
    info('getDocumentsByBatch fetched from dataverse', { batchId });
    return res.data.value;
  } catch (e: any) {
    const status = e?.response?.status;
    if (status === 404) {
      warn('getDocumentsByBatch: entity set not found', { set: DV_SETS.documentsSet });
      return [];
    }
    if (status === 401 || status === 403) {
      warn('getDocumentsByBatch: unauthorized/forbidden', { status });
      return [];
    }
    throw e;
  }
};

/** Fetch list of Businesses (for filtering/assignment). */
export const getBusinesses = async (token: string): Promise<Business[]> => {
  try {
    const res = await axios.get(`${baseUrl}/${DV_SETS.businessesSet}?$select=toba_businessid,toba_name,toba_code,toba_isactive`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="*"'
      }
    });
    info('getBusinesses fetched from dataverse', { count: Array.isArray(res.data?.value) ? res.data.value.length : 0 });
    return res.data.value;
  } catch (e: any) {
    const status = e?.response?.status;
    if (status === 404) {
      // Entity set not found in this environment; return empty for a graceful UI
      warn('getBusinesses: entity set not found', { set: DV_SETS.businessesSet });
      return [];
    }
    throw e;
  }
};

/** List Batch Recipients with optional filters. */
export const getBatchRecipients = async (
  token: string,
  filters?: { businessId?: string; department?: string; primaryGroup?: string }
): Promise<any[]> => {
  const qs: string[] = [
    `$select=toba_batchrecipientid,toba_name,${DV_ATTRS.batchRecipientUserField},${DV_ATTRS.batchRecipientEmailField},${DV_ATTRS.batchRecipientDisplayNameField},${DV_ATTRS.batchRecipientDepartmentField},${DV_ATTRS.batchRecipientJobTitleField},${DV_ATTRS.batchRecipientLocationField},${DV_ATTRS.batchRecipientPrimaryGroupField},_toba_batch_value,_toba_business_value`
  ];
  const filt: string[] = [];
  if (filters?.businessId) filt.push(`_toba_business_value eq ${filters.businessId.replace(/[{}]/g,'')}`);
  if (filters?.department) filt.push(`${DV_ATTRS.batchRecipientDepartmentField} eq '${String(filters.department).replace(/'/g, "''")}'`);
  if (filters?.primaryGroup) filt.push(`${DV_ATTRS.batchRecipientPrimaryGroupField} eq '${String(filters.primaryGroup).replace(/'/g, "''")}'`);
  if (filt.length) qs.push(`$filter=${filt.join(' and ')}`);
  const url = `${baseUrl}/${DV_SETS.batchRecipientsSet}?${qs.join('&')}`;
  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="*"'
    }
  });
  return Array.isArray(res.data?.value) ? res.data.value : [];
};

/**
 * Get acknowledged document IDs for a specific user within a batch.
 * Uses the acknowledgements table and returns document GUID strings.
 */
export const getAckedDocIdsForUser = async (batchId: string, token: string, userEmail?: string, userId?: string): Promise<string[]> => {
  const bid = (batchId || '').replace(/[{}]/g, '');
  const emailFilter = userEmail ? ` and ${DV_ATTRS.ackUserField} eq '${String(userEmail).toLowerCase().replace(/'/g, "''")}'` : '';
  const idFilter = (!emailFilter && userId) ? ` and _toba_user_value eq ${userId.replace(/[{}]/g, '')}` : '';
  const url = `${baseUrl}/${DV_SETS.userAcksSet}?$select=_toba_document_value&$filter=_toba_batch_value eq ${bid} and toba_acknowledged eq true${emailFilter || idFilter}`;
  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="*"'
    }
  });
  const rows = Array.isArray(res.data?.value) ? res.data.value : [];
  // _toba_document_value returns GUID without braces
  return rows.map((r: any) => r?._toba_document_value).filter(Boolean);
};

/**
 * Count records in a Dataverse set with an optional OData $filter.
 * Returns a numeric count using $count=true. When the API doesn’t return @odata.count,
 * falls back to reading the array length.
 */
export const countRecords = async (setName: string, token: string, filter?: string): Promise<number> => {
  const qs = [`$count=true`, `$select=${encodeURIComponent('toba_name')}`];
  if (filter) qs.push(`$filter=${filter}`);
  const url = `${baseUrl}/${setName}?${qs.join('&')}`;
  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      Prefer: 'odata.include-annotations="*",odata.maxpagesize=1'
    }
  });
  const count = (res.data && (res.data['@odata.count'] ?? res.data['@odata\.count'])) ?? (Array.isArray(res.data?.value) ? res.data.value.length : 0);
  info('dataverseService: countRecords', { set: setName, filter, count });
  return Number(count) || 0;
};

/** Generic Dataverse helpers for CRUD and metadata (live mode only). */
const orgBase = () => `${process.env.REACT_APP_DATAVERSE_URL?.replace(/\/$/, '')}/api/data/v9.2`;

export async function listRecords(
  setName: string,
  token: string,
  options?: { top?: number; select?: string[]; filter?: string; orderby?: string }
): Promise<any[]> {
  const qs: string[] = [];
  const top = options?.top ?? 50; if (top) qs.push(`$top=${top}`);
  const sel = options?.select && options.select.length ? `$select=${options.select.join(',')}` : undefined; if (sel) qs.push(sel);
  if (options?.filter) qs.push(`$filter=${options.filter}`);
  if (options?.orderby) qs.push(`$orderby=${options.orderby}`);
  const url = `${orgBase()}/${setName}${qs.length ? '?' + qs.join('&') : ''}`;
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'odata.include-annotations="*"' }
  });
  return Array.isArray(res.data?.value) ? res.data.value : [];
}

export async function getEntityLogicalNameBySet(setName: string, token: string): Promise<string | undefined> {
  const url = `${orgBase()}/EntityDefinitions?$select=EntitySetName,LogicalName&$filter=EntitySetName eq '${setName}'`;
  const res = await axios.get(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  const row = Array.isArray(res.data?.value) ? res.data.value[0] : undefined;
  return row?.LogicalName;
}

export async function getEntityAttributes(logicalName: string, token: string): Promise<string[]> {
  const url = `${orgBase()}/EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName&$top=500`;
  const res = await axios.get(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
  const vals = Array.isArray(res.data?.value) ? res.data.value.map((a: any) => a?.LogicalName).filter(Boolean) : [];
  return vals;
}

export async function createRecord(setName: string, token: string, body: any): Promise<{ id?: string; ok: boolean; status?: number; text?: string }> {
  const url = `${orgBase()}/${setName}`;
  try {
    const res = await axios.post(url, body, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'Content-Type': 'application/json' } });
    const loc = res.headers['odata-entityid'] || res.headers['OData-EntityId'] || '';
    const m = typeof loc === 'string' ? loc.match(/[0-9a-fA-F-]{36}/) : null;
    return { ok: true, id: m ? m[0] : undefined };
  } catch (e: any) {
    return { ok: false, status: e?.response?.status, text: typeof e?.response?.data === 'string' ? e.response.data : JSON.stringify(e?.response?.data || e.message) };
  }
}

export async function updateRecord(setName: string, id: string, token: string, body: any): Promise<{ ok: boolean; status?: number; text?: string }> {
  const url = `${orgBase()}/${setName}(${id.replace(/[{}]/g, '')})`;
  try {
    await axios.patch(url, body, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'Content-Type': 'application/json', 'If-Match': '*' } });
    return { ok: true };
  } catch (e: any) {
    return { ok: false, status: e?.response?.status, text: typeof e?.response?.data === 'string' ? e.response.data : JSON.stringify(e?.response?.data || e.message) };
  }
}

export async function deleteRecord(setName: string, id: string, token: string): Promise<{ ok: boolean; status?: number; text?: string }> {
  const url = `${orgBase()}/${setName}(${id.replace(/[{}]/g, '')})`;
  try {
    await axios.delete(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' } });
    return { ok: true };
  } catch (e: any) {
    return { ok: false, status: e?.response?.status, text: typeof e?.response?.data === 'string' ? e.response.data : JSON.stringify(e?.response?.data || e.message) };
  }
}

/**
 * Probe read access for a Dataverse entity set using a harmless $top=1 query (no writes).
 * Returns an object indicating whether the call succeeded and any HTTP status on failure.
 */
export const probeReadAccess = async (
  setName: string,
  token: string
): Promise<{ set: string; ok: boolean; status?: number; count?: number; error?: string }> => {
  const url = `${baseUrl}/${setName}?$top=1`;
  try {
    const res = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        Prefer: 'odata.include-annotations="*"'
      }
    });
    const arr = Array.isArray(res.data?.value) ? res.data.value : [];
    info('dataverseService: probeReadAccess OK', { set: setName, count: arr.length });
    return { set: setName, ok: true, count: arr.length };
  } catch (e: any) {
    const status = e?.response?.status;
    const msg = typeof e?.message === 'string' ? e.message : undefined;
    warn('dataverseService: probeReadAccess failed', { set: setName, status, error: msg });
    return { set: setName, ok: false, status, error: msg };
  }
};

/**
 * Compute a user's progress for a batch from Dataverse.
 * Strategy (resilient):
 * 1) Count total documents in the batch (toba_documents filtered by _toba_batch_value)
 * 2) Try to read a pre-computed progress row (toba_userprogresses). If unavailable, fall back to counting user acknowledgements (toba_useracknowledgements).
 *    Both table/column names are expected to be aligned with your Dataverse solution — adjust as needed.
 *
 * On any failure, this returns zeros and logs details for developers. The UI will show neutral progress while a toast informs users if configured by callers.
 */
export const getUserProgress = async (batchId: string, token?: string, userId?: string, userEmail?: string) => {
  const bid = (batchId || '').replace(/[{}]/g, '');
  try {
    if (!process.env.REACT_APP_DATAVERSE_URL) {
      warn('Dataverse URL not configured');
      return { acknowledged: 0, total: 0, percent: 0 };
    }
    if (!token) {
      warn('getUserProgress called without token');
    }

    // 1) Total documents in batch
    let total = 0;
    try {
      const docsRes = await axios.get(`${baseUrl}/${DV_SETS.documentsSet}?$select=toba_documentid&$filter=_toba_batch_value eq ${bid}`,
        { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'odata.include-annotations="*"' } });
      total = Array.isArray(docsRes.data?.value) ? docsRes.data.value.length : 0;
    } catch (e: any) {
      logError('getUserProgress: failed to fetch documents for total', {
        batchId, url: `${baseUrl}/toba_documents`, status: e?.response?.status, data: e?.response?.data
      });
      // keep going; total will be 0
    }

    // Short-circuit: no documents means zero progress
    if (total === 0) {
      info('getUserProgress: no documents found for batch, returning 0s', { batchId });
      return { acknowledged: 0, total: 0, percent: 0 };
    }

    // 2a) Attempt reading from a pre-computed progress table
    try {
      // If your solution stores one row per user+batch, filter by user as well.
      // Common patterns: _toba_batch_value eq {bid} and _toba_user_value eq {userId}
      const emailFilter = userEmail ? ` and ${DV_ATTRS.ackUserField} eq '${String(userEmail).toLowerCase().replace(/'/g, "''")}'` : '';
      const idFilter = (!emailFilter && userId) ? ` and _toba_user_value eq ${userId.replace(/[{}]/g, '')}` : '';
  const prUrl = `${baseUrl}/${DV_SETS.userProgressesSet}?$select=toba_acknowledged,toba_totaldocs&$filter=_toba_batch_value eq ${bid}${emailFilter || idFilter}`;
  const prRes = await axios.get(prUrl, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'odata.include-annotations="*"' } });
      const row = prRes.data?.value?.[0];
      if (row) {
        const acknowledged = Number(row.toba_acknowledged ?? row.toba_acknowledgedcount ?? 0);
        const t = Number(row.toba_totaldocs ?? row.toba_total ?? total);
        const percent = t === 0 ? 0 : Math.round((acknowledged / t) * 100);
        info('getUserProgress: used toba_userprogresses', { batchId, acknowledged, total: t, percent });
        return { acknowledged, total: t, percent };
      }
      // If no row, fall through to ack counting
    } catch (e: any) {
      // Table may not exist yet — continue to fallback
      warn('getUserProgress: userprogress table not available, falling back to acknowledgements', {
        batchId, status: e?.response?.status
      });
    }

    // 2b) Count acknowledgements directly
    try {
      // Expected columns: toba_acknowledged (bool), _toba_batch_value (lookup), _toba_user_value (lookup)
      const emailFilter = userEmail ? ` and ${DV_ATTRS.ackUserField} eq '${String(userEmail).toLowerCase().replace(/'/g, "''")}'` : '';
      const idFilter = (!emailFilter && userId) ? ` and _toba_user_value eq ${userId.replace(/[{}]/g, '')}` : '';
  const ackUrl = `${baseUrl}/${DV_SETS.userAcksSet}?$select=toba_useracknowledgementid&$filter=_toba_batch_value eq ${bid} and toba_acknowledged eq true${emailFilter || idFilter}`;
  const ackRes = await axios.get(ackUrl, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0', Prefer: 'odata.include-annotations="*"' } });
      const acknowledged = Array.isArray(ackRes.data?.value) ? ackRes.data.value.length : 0;
      const percent = total === 0 ? 0 : Math.round((acknowledged / total) * 100);
      info('getUserProgress: counted acknowledgements', { batchId, acknowledged, total, percent });
      return { acknowledged, total, percent };
    } catch (e: any) {
      logError('getUserProgress: failed to count acknowledgements', {
        batchId, status: e?.response?.status, data: e?.response?.data
      });
    }

    // Fallback: return zeros if everything failed
    return { acknowledged: 0, total, percent: 0 };
  } catch (e: any) {
    logError('getUserProgress: unexpected error', e);
    try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: 'Progress is unavailable right now' } })); } catch {}
    return { acknowledged: 0, total: 0, percent: 0 };
  }
};
