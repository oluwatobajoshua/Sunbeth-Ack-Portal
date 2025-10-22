import type { Batch, Doc } from '../types/models';
import { info, warn } from '../diagnostics/logger';

const KEY_BATCHES = 'mock_batches';
const KEY_DOCS = 'mock_docs_by_batch';
const KEY_ACKS = 'mock_user_acks';

const defaultBatches: Batch[] = [
  { toba_batchid: '1', toba_name: 'Q4 2025 — Code of Conduct', toba_startdate: '2025-10-01', toba_duedate: '2025-10-31', toba_status: 'inprogress' },
  { toba_batchid: '2', toba_name: 'Q3 2025 — Health & Safety', toba_startdate: '2025-07-01', toba_duedate: '2025-07-31', toba_status: 'completed' },
  { toba_batchid: '3', toba_name: 'IT Security & Privacy', toba_startdate: '2025-04-15', toba_duedate: '2025-05-15', toba_status: 'inprogress' }
];

const defaultDocs: Record<string, Doc[]> = {
  '1': [
    { toba_documentid: 'd1', toba_title: 'Code of Conduct', toba_version: 'v1.0', toba_requiressignature: false, toba_fileurl: '' },
    { toba_documentid: 'd2', toba_title: 'Anti-Bribery Policy', toba_version: 'v1.1', toba_requiressignature: false, toba_fileurl: '' },
    { toba_documentid: 'd3', toba_title: 'Whistleblower Policy', toba_version: 'v2.0', toba_requiressignature: true, toba_fileurl: '' }
  ],
  '2': [
    { toba_documentid: 'd4', toba_title: 'Health & Safety Guide', toba_version: 'v2.0', toba_requiressignature: false, toba_fileurl: '' }
  ],
  '3': [
    { toba_documentid: 'd5', toba_title: 'Password Policy', toba_version: 'v3.2', toba_requiressignature: true, toba_fileurl: '' },
    { toba_documentid: 'd6', toba_title: 'Email Usage Policy', toba_version: 'v1.4', toba_requiressignature: false, toba_fileurl: '' }
  ]
};

export const seed = (batches: Batch[] = defaultBatches, docs: Record<string, Doc[]> = defaultDocs, acks: Record<string, string[]> = {}) => {
  localStorage.setItem(KEY_BATCHES, JSON.stringify(batches));
  localStorage.setItem(KEY_DOCS, JSON.stringify(docs));
  localStorage.setItem(KEY_ACKS, JSON.stringify(acks));
  try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch (e) {}
  info('mockDatabaseService: seeded mock data');
};

export const clear = () => {
  localStorage.removeItem(KEY_BATCHES);
  localStorage.removeItem(KEY_DOCS);
  localStorage.removeItem(KEY_ACKS);
  try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch (e) {}
  info('mockDatabaseService: cleared mock data');
};

export const getBatches = async (): Promise<Batch[]> => {
  await new Promise(r => setTimeout(r, 120));
  try {
    const raw = localStorage.getItem(KEY_BATCHES);
    if (!raw) {
      // auto-seed defaults for convenience
      localStorage.setItem(KEY_BATCHES, JSON.stringify(defaultBatches));
      localStorage.setItem(KEY_DOCS, JSON.stringify(defaultDocs));
      if (!localStorage.getItem(KEY_ACKS)) localStorage.setItem(KEY_ACKS, JSON.stringify({}));
      try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {}
      return defaultBatches;
    }
    return JSON.parse(raw) as Batch[];
  } catch (e) {
    warn('mockDatabaseService: failed to parse batches, returning defaults');
    return defaultBatches;
  }
};

export const getDocumentsByBatch = async (batchId: string): Promise<Doc[]> => {
  await new Promise(r => setTimeout(r, 80));
  try {
    const raw = localStorage.getItem(KEY_DOCS);
    if (!raw) {
      // ensure defaults are present
      localStorage.setItem(KEY_DOCS, JSON.stringify(defaultDocs));
      try { window.dispatchEvent(new CustomEvent('mockDataChanged')); } catch {}
      return (defaultDocs[batchId] || []);
    }
    const map = JSON.parse(raw) as Record<string, Doc[]>;
    return map[batchId] || [];
  } catch (e) {
    warn('mockDatabaseService: failed to parse docs, returning defaults');
    return (defaultDocs[batchId] || []);
  }
};

export const getUserProgress = async (batchId: string) => {
  await new Promise(r => setTimeout(r, 60));
  try {
    const raw = localStorage.getItem(KEY_ACKS);
    const map: Record<string, string[]> = raw ? JSON.parse(raw) : {};
    const acked = map[batchId] || [];
    const docs = await getDocumentsByBatch(batchId);
    const total = docs.length;
    const percent = total === 0 ? 0 : Math.round((acked.length / total) * 100);
    info('mockDatabaseService: computed progress', { batchId, acknowledged: acked.length, total, percent });
    return { acknowledged: acked.length, total, percent };
  } catch (e) {
    warn('mockDatabaseService: error computing progress', e);
    return { acknowledged: 0, total: 0, percent: 0 };
  }
};

/** Return acknowledged document IDs for a batch (mock mode). */
export const getAcknowledgedDocIds = async (batchId: string): Promise<string[]> => {
  await new Promise(r => setTimeout(r, 40));
  try {
    const raw = localStorage.getItem(KEY_ACKS);
    const map: Record<string, string[]> = raw ? JSON.parse(raw) : {};
    return map[batchId] || [];
  } catch (e) {
    return [];
  }
};
