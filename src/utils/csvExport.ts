const esc = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
const toCsv = (rows: any[], headers: string[]): string => {
  const head = headers.map(esc).join(',');
  const body = rows.map(r => headers.map(h => esc(r[h])).join(',')).join('\n');
  return head + '\n' + body;
};

const triggerDownload = (filename: string, content: string, mime = 'text/csv;charset=utf-8;') => {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = filename; a.click(); URL.revokeObjectURL(url);
};

type ExportOpts = { year?: number | string, adminEmail?: string };

/**
 * Export core datasets to CSV files mirroring the Excel export and include acknowledgements:
 * - batches.csv
 * - documents.csv
 * - recipients.csv
 * - acknowledgements.csv (yearly, legal consent timestamp included)
 */
export const exportAnalyticsCsvFull = async (opts: ExportOpts = {}) => {
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  if (!sqliteEnabled) {
    throw new Error('SQLite API must be enabled to export analytics');
  }
  const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');
  const year = String((opts.year ?? new Date().getFullYear()));
  const adminEmail = String(opts.adminEmail || '').trim().toLowerCase();

  const [batchesRes, recRes] = await Promise.all([
    fetch(`${base}/api/batches`),
    fetch(`${base}/api/recipients`)
  ]);
  const batches = await batchesRes.json().catch(() => []);
  const recipients = await recRes.json().catch(() => []);

  const docsRows: any[] = [];
  for (const b of (Array.isArray(batches) ? batches : [])) {
    const id = String((b.toba_batchid || b.id));
    if (!id) continue;
    try {
      const dRes = await fetch(`${base}/api/batches/${encodeURIComponent(id)}/documents`);
      const rows = await dRes.json();
      for (const d of (Array.isArray(rows) ? rows : [])) {
        docsRows.push({ batchId: id, ...d });
      }
    } catch {}
  }

  const batchRows = (Array.isArray(batches) ? batches : []).map((r: any) => ({
    id: String(r.toba_batchid || r.id),
    name: String(r.toba_name || r.name || ''),
    startDate: r.toba_startdate || r.startDate || '',
    dueDate: r.toba_duedate || r.dueDate || '',
    status: r.toba_status != null ? String(r.toba_status) : (r.status != null ? String(r.status) : ''),
    description: r.description || ''
  }));

  const docRows = docsRows.map((d: any) => ({
    id: String(d.toba_documentid || d.id || ''),
    batchId: String(d.batchId || d.toba_batchid || ''),
    title: String(d.toba_title || d.title || ''),
    url: String(d.toba_fileurl || d.url || ''),
    version: Number(d.toba_version || d.version || 1),
    requiresSignature: (d.toba_requiressignature ?? d.requiresSignature) ? true : false,
    driveId: d.toba_driveid || d.driveId || '',
    itemId: d.toba_itemid || d.itemId || '',
    source: d.toba_source || d.source || ''
  }));

  const recRows = (Array.isArray(recipients) ? recipients : []).map((r: any) => ({
    id: Number(r.id),
    batchId: Number(r.batchId),
    businessId: r.businessId != null ? Number(r.businessId) : '',
    email: String(r.email || r.user || ''),
    displayName: String(r.displayName || ''),
    department: r.department || '',
    jobTitle: r.jobTitle || '',
    location: r.location || '',
    primaryGroup: r.primaryGroup || ''
  }));

  const batchesCsv = toCsv(batchRows, ['id','name','startDate','dueDate','status','description']);
  const documentsCsv = toCsv(docRows, ['id','batchId','title','url','version','requiresSignature','driveId','itemId','source']);
  const recipientsCsv = toCsv(recRows, ['id','batchId','businessId','email','displayName','department','jobTitle','location','primaryGroup']);

  triggerDownload(`batches-${year}.csv`, batchesCsv);
  triggerDownload(`documents-${year}.csv`, documentsCsv);
  triggerDownload(`recipients-${year}.csv`, recipientsCsv);

  // Acknowledgements
  try {
    const url = `${base}/api/admin/acks/export?year=${encodeURIComponent(year)}`;
    const headers: any = {};
    if (adminEmail) headers['x-admin-email'] = adminEmail;
    const res = await fetch(url, { headers });
    if (res.ok) {
      const json = await res.json().catch(() => ({} as any));
      const rows = Array.isArray(json.records) ? json.records : [];
      if (rows.length > 0) {
        const headersList = ['year','batchId','batchName','documentId','documentTitle','email','displayName','department','jobTitle','location','primaryGroup','businessId','acknowledgedAt','legalConsentedAt'];
        const csv = toCsv(rows.map((r: any) => ({
          year: String(r.year || year),
          batchId: String(r.batchId || ''),
          batchName: String(r.batchName || ''),
          documentId: String(r.documentId || ''),
          documentTitle: String(r.documentTitle || ''),
          email: String(r.email || ''),
          displayName: String(r.displayName || ''),
          department: r.department || '',
          jobTitle: r.jobTitle || '',
          location: r.location || '',
          primaryGroup: r.primaryGroup || '',
          businessId: r.businessId != null ? Number(r.businessId) : '',
          acknowledgedAt: String(r.acknowledgedAt || ''),
          legalConsentedAt: String(r.legalConsentedAt || '')
        })), headersList);
        triggerDownload(`acknowledgements-${year}.csv`, csv);
      }
    }
  } catch (e) {
    console.warn('Acknowledgements export failed or unavailable', e);
  }
};
