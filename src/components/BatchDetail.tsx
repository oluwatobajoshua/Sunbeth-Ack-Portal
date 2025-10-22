/**
 * BatchDetail: Lists documents for a selected batch.
 *
 * Loads from the configured backend (SQLite API or SharePoint Lists).
 * No artificial fallbacks; empty/error state is shown if the call fails.
 */
import React, { useEffect, useState } from 'react';
import { Link, useParams } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { getDocumentsByBatch, getAcknowledgedDocIds } from '../services/dbService';
import type { Doc } from '../types/models';

const BatchDetail: React.FC = () => {
  const { id } = useParams();
  const { token, account } = useAuth();
  const [docs, setDocs] = useState<Doc[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [ackIds, setAckIds] = useState<string[]>([]);

  useEffect(() => {
    if (!id) return;
    const run = async () => {
      try {
        setLoading(true);
        const list = await getDocumentsByBatch(id, token ?? undefined);
        setDocs(list);
        setError(null);
        // fetch acknowledged doc ids for current user
        const acks = await getAcknowledgedDocIds(id, token ?? undefined, account?.username || undefined);
        setAckIds(acks);
      } catch {
        setDocs([]);
        setError('Unable to load documents for this batch.');
      } finally { setLoading(false); }
    };
    run();
  }, [token, id, account?.username]);

  return (
    <div className="container">
      <div className="card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div className="title">Batch</div>
            <div className="muted small">Documents assigned</div>
          </div>
          <Link to="/"><button className="btn ghost">← Back</button></Link>
        </div>
        <hr style={{ margin: '12px 0', border: 'none', borderTop: '1px solid #f4f4f4' }} />

        {loading ? (
          <div className="doc-list">
            <div className="doc-row">
              <div className="doc-meta">
                <div className="skeleton circle" />
                <div style={{ display: 'grid', gap: 8 }}>
                  <div className="skeleton line" style={{ width: 220 }} />
                  <div className="skeleton line" style={{ width: 140 }} />
                </div>
              </div>
              <div className="skeleton line" style={{ width: 80 }} />
            </div>
            <div className="doc-row">
              <div className="doc-meta">
                <div className="skeleton circle" />
                <div style={{ display: 'grid', gap: 8 }}>
                  <div className="skeleton line" style={{ width: 240 }} />
                  <div className="skeleton line" style={{ width: 120 }} />
                </div>
              </div>
              <div className="skeleton line" style={{ width: 80 }} />
            </div>
          </div>
        ) : docs.length === 0 ? (
          <div style={{ padding: 12 }}>
            <div className="muted">{error ? error : 'No documents found.'}</div>
          </div>
        ) : (
          <div className="doc-list">
            {docs.map((d: Doc, i: number) => (
              <div key={d.toba_documentid} className="doc-row">
                <div className="doc-meta">
                  <div className="doc-icon">PDF</div>
                  <div>
                    <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 8 }}>
                      <span>{i + 1}. {d.toba_title}</span>
                      {ackIds.includes(d.toba_documentid) && <span className="badge done">Acknowledged</span>}
                    </div>
                    <div className="muted small">{d.toba_version || ''}</div>
                  </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <Link to={`/document/${d.toba_documentid}?batchId=${id}`}><button className="btn sm">{ackIds.includes(d.toba_documentid) ? 'View' : 'Read'}</button></Link>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
export default BatchDetail;
