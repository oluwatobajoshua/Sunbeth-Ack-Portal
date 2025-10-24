import * as XLSX from 'xlsx';

const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
};

const toCsvBlob = (rows: Record<string, any>[]) => {
  if (!rows || rows.length === 0) return new Blob([''], { type: 'text/csv;charset=utf-8' });
  const headers = Object.keys(rows[0]);
  const lines = [headers.join(',')];
  for (const r of rows) {
    const line = headers.map(h => {
      const v = r[h] == null ? '' : String(r[h]);
      const needsQuote = /[",\n]/.test(v);
      const safe = v.replace(/"/g, '""');
      return needsQuote ? `"${safe}"` : safe;
    }).join(',');
    lines.push(line);
  }
  return new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
};

export const downloadExternalUsersTemplateExcel = () => {
  const wb = XLSX.utils.book_new();
  const readme = [
    { Field: 'email', Required: 'Yes', Notes: 'Primary key. Valid email address.' },
    { Field: 'name', Required: 'No', Notes: 'Full name (optional).' },
    { Field: 'phone', Required: 'No', Notes: 'Phone number (optional).' },
    { Field: 'status', Required: 'No', Notes: 'invited | active | disabled (default: invited for new records)' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(readme), 'READ_ME');
  const rows = [
    { email: 'user1@example.com', name: 'Ada Lovelace', phone: '+2348100000000', status: 'invited' },
    { email: 'user2@example.com', name: 'Grace Hopper', phone: '+2348100000001', status: 'active' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'ExternalUsers');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  downloadBlob(new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'external-users-template.xlsx');
};

export const downloadExternalUsersTemplateCsv = () => {
  const rows = [
    { email: 'user1@example.com', name: 'Ada Lovelace', phone: '+2348100000000', status: 'invited' },
    { email: 'user2@example.com', name: 'Grace Hopper', phone: '+2348100000001', status: 'active' },
  ];
  downloadBlob(toCsvBlob(rows), 'external-users-template.csv');
};

export const downloadBusinessesTemplateExcel = () => {
  const wb = XLSX.utils.book_new();
  const readme = [
    { Field: 'name', Required: 'Yes', Notes: 'Business display name. Used for upsert if code is missing.' },
    { Field: 'code', Required: 'No', Notes: 'Unique short code (preferred upsert key).' },
    { Field: 'description', Required: 'No', Notes: 'Optional description.' },
    { Field: 'isActive', Required: 'No', Notes: 'true|false or 1|0 (default: true for new records)' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(readme), 'READ_ME');
  const rows = [
    { name: 'Retail', code: 'RET', description: 'Retail business unit', isActive: true },
    { name: 'Wholesale', code: 'WHO', description: 'Wholesale unit', isActive: 1 },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Businesses');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  downloadBlob(new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'businesses-template.xlsx');
};

export const downloadBusinessesTemplateCsv = () => {
  const rows = [
    { name: 'Retail', code: 'RET', description: 'Retail business unit', isActive: true },
    { name: 'Wholesale', code: 'WHO', description: 'Wholesale unit', isActive: 1 },
  ];
  downloadBlob(toCsvBlob(rows), 'businesses-template.csv');
};

export const downloadAllTemplatesExcel = () => {
  const wb = XLSX.utils.book_new();
  const readme = [
    { Section: 'ExternalUsers', Notes: 'email (required), name, phone, status: invited|active|disabled' },
    { Section: 'Businesses', Notes: 'name (required), code (preferred), description, isActive: true|false|1|0' },
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(readme), 'README');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { email: 'user1@example.com', name: 'Ada Lovelace', phone: '+2348100000000', status: 'invited' }
  ]), 'ExternalUsers');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { name: 'Retail', code: 'RET', description: 'Retail business unit', isActive: true }
  ]), 'Businesses');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  downloadBlob(new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'import-templates.xlsx');
};
