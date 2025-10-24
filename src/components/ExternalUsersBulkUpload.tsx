import React, { useState } from 'react';
import { getApiBase } from '../utils/runtimeConfig';
import { showToast } from '../utils/alerts';

const ExternalUsersBulkUpload: React.FC = () => {
	const base = (getApiBase() as string) || '';
	const [uploading, setUploading] = useState(false);

	const onChange = async (file: File | null) => {
		if (!file) return;
		setUploading(true);
		try {
			const fd = new FormData();
			fd.append('file', file, file.name);
			const res = await fetch(`${base}/api/external-users/bulk-upload`, { method: 'POST', body: fd });
			const j = await res.json().catch(() => ({}));
			if (!res.ok) throw new Error(j?.error || 'bulk_upload_failed');
			showToast(`Bulk uploaded: ${j?.inserted || 0} inserted, ${j?.updated || 0} updated`, 'success');
		} catch (e) {
			showToast('Bulk upload failed', 'error');
		} finally { setUploading(false); }
	};

	return (
		<div className="card" style={{ padding: 12 }}>
			<div style={{ fontWeight: 700, marginBottom: 8 }}>External Users Bulk Upload</div>
			<input type="file" accept=".csv,.xlsx,.xls" onChange={e => onChange(e.target.files?.[0] || null)} disabled={uploading} />
			{uploading && <div className="small muted" style={{ marginTop: 6 }}>Uploading...</div>}
		</div>
	);
};

export default ExternalUsersBulkUpload;
