import * as XLSX from 'xlsx';

/**
 * Export core datasets to an Excel workbook with multiple sheets for analysis.
 * Sheets: Batches, Documents, Recipients
 */
export const exportAnalyticsExcel = async () => {
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  if (!sqliteEnabled) {
    throw new Error('SQLite API must be enabled to export analytics');
  }
  const base = (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '');

  // Fetch batches and recipients
  const [batchesRes, recRes] = await Promise.all([
    fetch(`${base}/api/batches`),
    fetch(`${base}/api/recipients`)
  ]);
  const batches = await batchesRes.json().catch(() => []);
  const recipients = await recRes.json().catch(() => []);

  // For documents, we need to query per batch to align with server contract
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

  // Map/normalize columns
  const batchSheetRows = (Array.isArray(batches) ? batches : []).map((r: any) => ({
    id: String(r.toba_batchid || r.id),
    name: String(r.toba_name || r.name || ''),
    startDate: r.toba_startdate || r.startDate || '',
    dueDate: r.toba_duedate || r.dueDate || '',
    status: r.toba_status != null ? String(r.toba_status) : (r.status != null ? String(r.status) : ''),
    description: r.description || ''
  }));

  const docSheetRows = docsRows.map((d: any) => ({
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

  const recSheetRows = (Array.isArray(recipients) ? recipients : []).map((r: any) => ({
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

  const wb = XLSX.utils.book_new();
  const wsBatches = XLSX.utils.json_to_sheet(batchSheetRows);
  const wsDocs = XLSX.utils.json_to_sheet(docSheetRows);
  const wsRecs = XLSX.utils.json_to_sheet(recSheetRows);

  XLSX.utils.book_append_sheet(wb, wsBatches, 'Batches');
  XLSX.utils.book_append_sheet(wb, wsDocs, 'Documents');
  XLSX.utils.book_append_sheet(wb, wsRecs, 'Recipients');

  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'sunbeth-ack-analytics.xlsx';
  a.click();
  URL.revokeObjectURL(url);
};
