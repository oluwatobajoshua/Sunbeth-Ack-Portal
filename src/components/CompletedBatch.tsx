import React, { useEffect, useMemo, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { getDocumentsByBatch, getAcknowledgedDocIds } from '../services/dbService';
import type { Doc } from '../types/models';

const CompletedBatch: React.FC = () => {
  const { id } = useParams();
  const { token, account } = useAuth();
  const [docs, setDocs] = useState<Doc[]>([]);
  const [ackIds, setAckIds] = useState<string[]>([]);

  useEffect(() => {
    if (!id) return;
    (async () => {
      const list = await getDocumentsByBatch(id, token ?? undefined);
      setDocs(list);
      const a = await getAcknowledgedDocIds(id, token ?? undefined, account?.username || undefined);
      setAckIds(a);
    })();
  }, [id, token, account?.username]);

  const ackedDocs = useMemo(() => docs.filter(d => ackIds.includes(d.toba_documentid)), [docs, ackIds]);

  return (
    <div className="container">
      <div className="card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div className="title">Completed Documents</div>
            <div className="muted small">You can still view previously acknowledged documents.</div>
          </div>
          <Link to="/"><button className="btn ghost sm">← Back to Dashboard</button></Link>
        </div>
        <hr style={{ margin: '12px 0', border: 'none', borderTop: '1px solid #f4f4f4' }} />

        {ackedDocs.length === 0 ? (
          <div className="muted" style={{ padding: 12 }}>No acknowledged documents found.</div>
        ) : (
          <div className="doc-list">
            {ackedDocs.map((d, i) => (
              <div key={d.toba_documentid} className="doc-row">
                <div className="doc-meta">
                  <div className="doc-icon">PDF</div>
                  <div>
                    <div style={{ fontWeight: 700 }}>{i + 1}. {d.toba_title}</div>
                    <div className="muted small">{d.toba_version || ''}</div>
                  </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  {/* open the same reader in view mode; we can add a view-only flag later */}
                  <Link to={`/document/${d.toba_documentid}?batchId=${id}`}><button className="btn sm">View</button></Link>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default CompletedBatch;
