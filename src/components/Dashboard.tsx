/**
 * Dashboard: Shows assigned batches with per-batch progress.
 *
 * - In live mode, data is fetched from Dataverse via dbService; on error an empty/error state is shown.
 * - In mock mode, the Dev Panel can seed local data; progress updates on 'mockAck' events.
 */
import React, { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { getBatches, getUserProgress } from '../services/dbService';
import type { Batch } from '../types/models';
import { useRuntimeMock } from '../utils/runtimeMock';

const Dashboard: React.FC = () => {
  const { token, account } = useAuth();
  const [batches, setBatches] = useState<Batch[]>([]);
  const [progressMap, setProgressMap] = useState<Record<string, { percent: number; total: number; acknowledged: number }>>({});
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const runtimeMock = useRuntimeMock();

  useEffect(() => {
    const load = async () => {
      try {
        setLoading(true);
  const list = await getBatches(runtimeMock ? undefined : token ?? undefined, account?.username || undefined);
        setBatches(Array.isArray(list) ? list : []);
        setError(null);
      } catch {
        setBatches([]);
        setError('Unable to load your batches.');
      } finally { setLoading(false); }
    };
    const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
    if (runtimeMock || sqliteEnabled) load();
    else if (token) load();
  }, [token, runtimeMock, account?.username]);

  useEffect(() => {
  if (!Array.isArray(batches) || batches.length === 0) return;
    // fetch progress per-batch (mock)
    (async () => {
      const m: Record<string, { percent: number; total: number; acknowledged: number }> = {};
      for (const b of batches) {
        try {
          const p = await getUserProgress(b.toba_batchid, token ?? undefined, undefined, account?.username || undefined);
          m[b.toba_batchid] = { percent: p.percent, total: p.total ?? 0, acknowledged: p.acknowledged ?? 0 };
        } catch {
          m[b.toba_batchid] = { percent: 0, total: 0, acknowledged: 0 };
        }
      }
      setProgressMap(m);
    })();
  }, [batches, token, account?.username]);

  // listen for mock ack events to refresh progress
  useEffect(() => {
    const h = async (e: Event) => {
      const ev = e as CustomEvent<any>;
      const batchId = ev?.detail?.batchId;
      if (!batchId) return;
      try {
        const p = await getUserProgress(batchId, token ?? undefined);
        setProgressMap(prev => ({ ...prev, [batchId]: { percent: p.percent, total: p.total ?? 0, acknowledged: p.acknowledged ?? 0 } }));
      } catch { }
    };
    window.addEventListener('mockAck', h as EventListener);
    window.addEventListener('sunbeth:progressUpdated', h as EventListener);
    return () => {
      window.removeEventListener('mockAck', h as EventListener);
      window.removeEventListener('sunbeth:progressUpdated', h as EventListener);
    };
  }, [runtimeMock, token]);

  const formatDate = (d?: string) => {
    if (!d) return '—';
    try { return new Date(d).toLocaleDateString(); } catch { return d; }
  };

  // Be defensive in case a test or edge case provides a non-array value
  const batchList = Array.isArray(batches) ? batches : ([] as Batch[]);
  const incompleteCount = useMemo(() => batchList.filter(b => (progressMap[b.toba_batchid]?.percent ?? 0) < 100).length, [batchList, progressMap]);
  const earliestDue = useMemo(() => {
    const incompletes = batchList.filter(b => (progressMap[b.toba_batchid]?.percent ?? 0) < 100);
    const dates = incompletes.map(b => b.toba_duedate).filter(Boolean) as string[];
    if (!dates.length) return null;
    const min = dates.reduce((a, d) => (new Date(d) < new Date(a) ? d! : a!));
    return min;
  }, [batchList, progressMap]);

  return (
    <div className="container">
      <div className="card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div className="title">Welcome{account ? `, ${account.name}` : ''}</div>
            <div className="muted" style={{ marginTop: 6 }}>You have <strong>{incompleteCount}</strong> pending items</div>
            {account?.username && (
              <div className="small muted" style={{ marginTop: 4 }}>Signed in as: {account.username}</div>
            )}
          </div>
          <div style={{ textAlign: 'right' }}>
            <div className="muted small">Due by: <strong style={{ color: 'var(--accent)' }}>{earliestDue ? formatDate(earliestDue) : '—'}</strong></div>
          </div>
        </div>
        <hr style={{ margin: '16px 0', border: 'none', borderTop: '1px solid #f2f2f2' }} />

        {loading ? (
          <div style={{ display: 'grid', gap: 10 }}>
            <div className="skeleton block" />
            <div className="skeleton block" />
            <div className="skeleton block" />
          </div>
        ) : batchList.length === 0 ? (
          <div style={{ padding: 12 }}>
            <div className="muted">{error ? error : 'No batches assigned.'}</div>
          </div>
        ) : (
          <div className="batch-list">
            {batchList.map(b => (
              <div key={b.toba_batchid} className="batch-tile">
                <div>
                  <div style={{ fontWeight: 700 }}>{b.toba_name}</div>
                  <div className="muted" style={{ marginTop: 6 }}>Mandatory for all staff • Due: {formatDate(b.toba_duedate)}</div>
                  <div style={{ marginTop: 8 }}>
                    <div className="progressBar" aria-hidden="true"><i style={{ width: `${progressMap[b.toba_batchid]?.percent || 0}%` }} /></div>
                    <div className="muted small" style={{ marginTop: 6 }}>{progressMap[b.toba_batchid]?.percent || 0}% acknowledged</div>
                    {(() => {
                      const prog = progressMap[b.toba_batchid] || { percent: 0, total: 0, acknowledged: 0 };
                      const remaining = Math.max(0, (prog.total ?? 0) - (prog.acknowledged ?? 0));
                      return (
                        <div className="stats-row" aria-label="batch document stats">
                          <span className="chip"><strong>{prog.total ?? 0}</strong> Docs</span>
                          <span className="chip ok"><strong>{prog.acknowledged ?? 0}</strong> Acknowledged</span>
                          <span className="chip warn"><strong>{remaining}</strong> Remaining</span>
                        </div>
                      );
                    })()}
                  </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div>{(progressMap[b.toba_batchid]?.percent || 0) === 100 ? <span className="badge done">Completed</span> : <span className="badge progress">In Progress</span>}</div>
                  <div style={{ marginTop: 10 }}>
                    { (progressMap[b.toba_batchid]?.percent || 0) === 100 ? (
                      <Link to={`/batch/${b.toba_batchid}/completed`}><button className="btn ghost sm">View</button></Link>
                    ) : (
                      <Link to={`/batch/${b.toba_batchid}`}><button className="btn ghost sm">Continue</button></Link>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
export default Dashboard;
