import React, { useState } from 'react';
import { getApiBase } from '../utils/runtimeConfig';
import { showToast } from '../utils/alerts';
import { downloadBusinessesTemplateExcel, downloadBusinessesTemplateCsv } from '../utils/importTemplates';

const BusinessesBulkUpload: React.FC = () => {
  const base = (getApiBase() as string) || '';
  const [uploading, setUploading] = useState(false);

  const onChange = async (file: File | null) => {
    if (!file) return;
    setUploading(true);
    try {
      const fd = new FormData();
      fd.append('file', file, file.name);
      const res = await fetch(`${base}/api/businesses/bulk-upload`, { method: 'POST', body: fd });
      const j = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(j?.error || 'bulk_upload_failed');
      showToast(`Processed: ${j?.inserted || 0} inserted, ${j?.updated || 0} updated`, 'success');
      if (Array.isArray(j?.errors) && j.errors.length > 0) {
        const headers = Object.keys(j.errors[0]);
        const lines = [headers.join(',')].concat(j.errors.map((r: any) => headers.map(h => {
          const v = r[h] == null ? '' : String(r[h]);
          const needsQuote = /[",\n]/.test(v);
          const safe = v.replace(/"/g, '""');
          return needsQuote ? `"${safe}"` : safe;
        }).join(',')));
        const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'businesses-errors.csv'; a.click(); URL.revokeObjectURL(url);
        showToast(`${j.errors.length} row(s) had issues. Downloaded error report.`, 'warning');
      }
    } catch (e) {
      showToast('Bulk upload failed', 'error');
    } finally { setUploading(false); }
  };

  return (
    <div className="card" style={{ padding: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
        <div>
          <div style={{ fontWeight: 700 }}>Businesses Bulk Upload</div>
          <div className="small muted">Upload Businesses from CSV or Excel.</div>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button className="btn ghost xs" onClick={downloadBusinessesTemplateExcel}>Template (Excel)</button>
          <button className="btn ghost xs" onClick={downloadBusinessesTemplateCsv}>Template (CSV)</button>
        </div>
      </div>
      <input type="file" accept=".csv,.xlsx,.xls" onChange={e => onChange(e.target.files?.[0] || null)} disabled={uploading} />
      {uploading && <div className="small muted" style={{ marginTop: 6 }}>Uploading...</div>}
    </div>
  );
};

export default BusinessesBulkUpload;
