import { getGraphToken } from './authTokens';

const GRAPH = 'https://graph.microsoft.com/v1.0';
const OVERRIDE_KEY = 'sunbeth:sp:siteIdOverride';

function getSiteId(): string {
  try {
    const ov = localStorage.getItem(OVERRIDE_KEY);
    if (ov && ov.trim()) return ov.trim();
  } catch {}
  return (process.env.REACT_APP_SP_SITE_ID || '').trim();
}

export function setSharePointSiteIdOverride(id: string | null) {
  try {
    if (!id) localStorage.removeItem(OVERRIDE_KEY);
    else localStorage.setItem(OVERRIDE_KEY, id);
  } catch {}
}

export const SP_LISTS = {
  batches: process.env.REACT_APP_SP_LIST_BATCHES || 'sunbeth_batches',
  documents: process.env.REACT_APP_SP_LIST_DOCUMENTS || 'sunbeth_documents',
  recipients: process.env.REACT_APP_SP_LIST_RECIPIENTS || 'sunbeth_recipients',
  acks: process.env.REACT_APP_SP_LIST_ACKS || 'sunbeth_acks',
  progress: process.env.REACT_APP_SP_LIST_PROGRESS || 'sunbeth_progress',
  businesses: process.env.REACT_APP_SP_LIST_BUSINESSES || 'sunbeth_businesses'
};

export type ProvisionStep = { step: string; ok: boolean; detail?: string };

async function gFetch(path: string, init?: RequestInit, scopes: string[] = ['Sites.ReadWrite.All']): Promise<Response> {
  const token = await getGraphToken(scopes);
  const res = await fetch(`${GRAPH}${path}`, {
    ...(init || {}),
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      ...(init?.headers || {})
    }
  });
  return res;
}

async function getListByName(name: string): Promise<{ id: string; name: string } | null> {
  const siteId = getSiteId();
  if (!siteId) throw new Error('REACT_APP_SP_SITE_ID not set');
  const res = await gFetch(`/sites/${siteId}/lists?$filter=displayName eq '${name.replace(/'/g, "''")}'`);
  if (!res.ok) throw new Error(`getListByName failed: ${res.status}`);
  const j = await res.json().catch(() => ({ value: [] }));
  const row = Array.isArray(j?.value) ? j.value[0] : null;
  return row ? { id: row.id, name: row.displayName } : null;
}

async function ensureList(name: string): Promise<{ id: string }> {
  const siteId = getSiteId();
  const existing = await getListByName(name).catch(() => null);
  if (existing) return { id: existing.id };
  const res = await gFetch(`/sites/${siteId}/lists`, {
    method: 'POST',
    body: JSON.stringify({ displayName: name, list: { template: 'genericList' } })
  });
  if (!res.ok) throw new Error(`Create list ${name} failed: ${res.status} ${await res.text().catch(() => '')}`);
  const j = await res.json().catch(() => ({}));
  return { id: j.id };
}

async function ensureColumn(listId: string, col: any): Promise<void> {
  // Try to create; ignore 409 if exists
  const siteId = getSiteId();
  const res = await gFetch(`/sites/${siteId}/lists/${listId}/columns`, { method: 'POST', body: JSON.stringify(col) });
  if (res.ok) return;
  if (res.status === 409) return; // already exists
  // Some tenants return 400 when exists; fetch to confirm
  if (res.status === 400) return;
  throw new Error(`Create column failed: ${res.status} ${await res.text().catch(() => '')}`);
}

export async function provisionSharePointLists(): Promise<ProvisionStep[]> {
  const logs: ProvisionStep[] = [];
  const siteId = getSiteId();
  if (!siteId) return [{ step: 'SharePoint Site', ok: false, detail: 'REACT_APP_SP_SITE_ID not set' }];
  try {
    const batches = await ensureList(SP_LISTS.batches); logs.push({ step: 'List batches', ok: true, detail: batches.id });
    const documents = await ensureList(SP_LISTS.documents); logs.push({ step: 'List documents', ok: true, detail: documents.id });
    const recipients = await ensureList(SP_LISTS.recipients); logs.push({ step: 'List recipients', ok: true, detail: recipients.id });
    const acks = await ensureList(SP_LISTS.acks); logs.push({ step: 'List acks', ok: true, detail: acks.id });
    const progress = await ensureList(SP_LISTS.progress); logs.push({ step: 'List progress', ok: true, detail: progress.id });
    const businesses = await ensureList(SP_LISTS.businesses); logs.push({ step: 'List businesses', ok: true, detail: businesses.id });

    // Add columns
    // Batches: Title (name) exists by default
    await ensureColumn(batches.id, { name: 'startDate', text: {} });
    await ensureColumn(batches.id, { name: 'dueDate', text: {} });
    await ensureColumn(batches.id, { name: 'status', number: {} });
    await ensureColumn(batches.id, { name: 'description', text: { allowMultipleLines: true } });

    // Documents
    await ensureColumn(documents.id, { name: 'batchId', number: {} });
    await ensureColumn(documents.id, { name: 'title', text: {} });
    await ensureColumn(documents.id, { name: 'url', hyperlink: {} });
    await ensureColumn(documents.id, { name: 'version', number: {} });
    await ensureColumn(documents.id, { name: 'requiresSignature', boolean: {} });

    // Recipients
    await ensureColumn(recipients.id, { name: 'batchId', number: {} });
    await ensureColumn(recipients.id, { name: 'businessId', number: {} });
    await ensureColumn(recipients.id, { name: 'user', text: {} });
    await ensureColumn(recipients.id, { name: 'email', text: {} });
    await ensureColumn(recipients.id, { name: 'displayName', text: {} });
    await ensureColumn(recipients.id, { name: 'department', text: {} });
    await ensureColumn(recipients.id, { name: 'jobTitle', text: {} });
    await ensureColumn(recipients.id, { name: 'location', text: {} });
    await ensureColumn(recipients.id, { name: 'primaryGroup', text: {} });

    // Acks
    await ensureColumn(acks.id, { name: 'batchId', number: {} });
    await ensureColumn(acks.id, { name: 'documentId', number: {} });
    await ensureColumn(acks.id, { name: 'user', text: {} });
    await ensureColumn(acks.id, { name: 'email', text: {} });
    await ensureColumn(acks.id, { name: 'acknowledged', boolean: {} });
    await ensureColumn(acks.id, { name: 'ackDate', text: {} });

    // Progress
    await ensureColumn(progress.id, { name: 'batchId', number: {} });
    await ensureColumn(progress.id, { name: 'user', text: {} });
    await ensureColumn(progress.id, { name: 'email', text: {} });
    await ensureColumn(progress.id, { name: 'acknowledged', number: {} });
    await ensureColumn(progress.id, { name: 'totalDocs', number: {} });

    // Businesses: Title as name
    await ensureColumn(businesses.id, { name: 'code', text: {} });
    await ensureColumn(businesses.id, { name: 'isActive', boolean: {} });
    await ensureColumn(businesses.id, { name: 'description', text: { allowMultipleLines: true } });

    return logs;
  } catch (e: any) {
    logs.push({ step: 'Provision failed', ok: false, detail: String(e?.message || e) });
    return logs;
  }
}

async function getItems(listName: string, filter?: string): Promise<any[]> {
  const list = await getListByName(listName);
  if (!list) return [];
  const siteId = getSiteId();
  const qs = filter ? `?$filter=${encodeURIComponent(filter)}` : '';
  const res = await gFetch(`/sites/${siteId}/lists/${list.id}/items${qs}?$expand=fields`);
  if (!res.ok) return [];
  const j = await res.json().catch(() => ({ value: [] }));
  return Array.isArray(j?.value) ? j.value.map((it: any) => ({ id: it.id, ...it.fields })) : [];
}

async function getItemById(listName: string, id: number): Promise<any | null> {
  const list = await getListByName(listName);
  if (!list) return null;
  const siteId = getSiteId();
  const res = await gFetch(`/sites/${siteId}/lists/${list.id}/items/${id}?$expand=fields`);
  if (!res.ok) return null;
  const j = await res.json().catch(() => null);
  return j ? { id: j.id, ...j.fields } : null;
}

async function addItem(listName: string, fields: any): Promise<number | null> {
  const list = await getListByName(listName);
  if (!list) return null;
  const siteId = getSiteId();
  const res = await gFetch(`/sites/${siteId}/lists/${list.id}/items`, { method: 'POST', body: JSON.stringify({ fields }) });
  if (!res.ok) return null;
  const j = await res.json().catch(() => null);
  return j?.id ?? null;
}

export async function spGetBusinesses(): Promise<any[]> {
  return getItems(SP_LISTS.businesses);
}

export async function spGetBatches(userEmail?: string): Promise<any[]> {
  if (!userEmail) return getItems(SP_LISTS.batches);
  const recips = await getItems(SP_LISTS.recipients, `fields/email eq '${userEmail.replace(/'/g, "''")}'`);
  const ids = Array.from(new Set(recips.map(r => Number(r.batchId)).filter(Boolean)));
  const results: any[] = [];
  for (const id of ids) {
    const row = await getItemById(SP_LISTS.batches, id).catch(() => null);
    if (row) results.push(row);
  }
  return results;
}

export async function spGetDocumentsByBatch(batchId: number): Promise<any[]> {
  return getItems(SP_LISTS.documents, `fields/batchId eq ${batchId}`);
}

export async function spGetBatchRecipients(filters?: { businessId?: number; department?: string; primaryGroup?: string }): Promise<any[]> {
  const f: string[] = [];
  if (filters?.businessId) f.push(`fields/businessId eq ${filters.businessId}`);
  if (filters?.department) f.push(`fields/department eq '${filters.department.replace(/'/g, "''")}'`);
  if (filters?.primaryGroup) f.push(`fields/primaryGroup eq '${filters.primaryGroup.replace(/'/g, "''")}'`);
  return getItems(SP_LISTS.recipients, f.length ? f.join(' and ') : undefined);
}

export async function spGetAckedDocIdsForUser(batchId: number, userEmail?: string): Promise<number[]> {
  const f: string[] = [`fields/batchId eq ${batchId}`, `fields/acknowledged eq true`];
  if (userEmail) f.push(`fields/email eq '${userEmail.replace(/'/g, "''")}'`);
  const rows = await getItems(SP_LISTS.acks, f.join(' and '));
  return rows.map(r => Number(r.documentId)).filter(Boolean);
}

export async function spGetUserProgress(batchId: number, userEmail?: string): Promise<{ acknowledged: number; total: number; percent: number }> {
  const docs = await spGetDocumentsByBatch(batchId);
  const total = docs.length;
  if (total === 0) return { acknowledged: 0, total: 0, percent: 0 };
  const acks = await getItems(SP_LISTS.acks, `fields/batchId eq ${batchId} and fields/acknowledged eq true${userEmail ? ` and fields/email eq '${userEmail.replace(/'/g, "''")}'` : ''}`);
  const acknowledged = acks.length;
  return { acknowledged, total, percent: total === 0 ? 0 : Math.round((acknowledged / total) * 100) };
}

// Creation helpers for Admin
export async function spCreateBatch(payload: { name: string; startDate?: string; dueDate?: string; description?: string; status?: number }): Promise<number | null> {
  return addItem(SP_LISTS.batches, { Title: payload.name, startDate: payload.startDate || null, dueDate: payload.dueDate || null, status: payload.status ?? 1, description: payload.description || null });
}

export async function spCreateDocument(batchId: number, d: { title: string; url: string; version?: number; requiresSignature?: boolean }): Promise<number | null> {
  return addItem(SP_LISTS.documents, { batchId, title: d.title, url: d.url, version: d.version ?? 1, requiresSignature: !!d.requiresSignature });
}

export async function spCreateRecipient(batchId: number, r: { businessId?: number | null; user?: string | null; email?: string; displayName?: string | null; department?: string | null; jobTitle?: string | null; location?: string | null; primaryGroup?: string | null }): Promise<number | null> {
  return addItem(SP_LISTS.recipients, { batchId, businessId: r.businessId ?? null, user: r.user || null, email: r.email || null, displayName: r.displayName || null, department: r.department || null, jobTitle: r.jobTitle || null, location: r.location || null, primaryGroup: r.primaryGroup || null });
}

// Discovery helpers
export async function findSitesByName(query: string): Promise<Array<{ id: string; displayName: string; webUrl: string }>> {
  const token = await getGraphToken(['Sites.Read.All']);
  const res = await fetch(`${GRAPH}/sites?search=${encodeURIComponent(query)}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return [];
  const j = await res.json().catch(() => ({ value: [] }));
  return (Array.isArray(j?.value) ? j.value : []).map((s: any) => ({ id: s.id, displayName: s.name || s.displayName, webUrl: s.webUrl })).filter((s: any) => s.id && (s.displayName || s.webUrl));
}
