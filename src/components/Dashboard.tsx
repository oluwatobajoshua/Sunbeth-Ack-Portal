/**
 * Dashboard: Shows assigned batches with per-batch progress.
 *
 * Data is fetched from the configured backend via dbService; on error an empty/error state is shown.
 */
import React, { useEffect, useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useRBAC } from '../context/RBACContext';
import { getBatches, getUserProgress } from '../services/dbService';
import type { Batch } from '../types/models';
import { fetchDuePolicies } from '../utils/policiesDue';

const Dashboard: React.FC = () => {
  const { token, account } = useAuth();
  const [batches, setBatches] = useState<Batch[]>([]);
  const [progressMap, setProgressMap] = useState<Record<string, { percent: number; total: number; acknowledged: number }>>({});
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [duePoliciesCount, setDuePoliciesCount] = useState<number>(0);
  const rbac = useRBAC();

  useEffect(() => {
    const load = async () => {
      try {
        setLoading(true);
        const list = await getBatches(token ?? undefined, account?.username || undefined);
        setBatches(Array.isArray(list) ? list : []);
        setError(null);
      } catch {
        setBatches([]);
        setError('Unable to load your batches.');
      } finally { setLoading(false); }
    };
    load();
  }, [token, account?.username]);

  // Pre-modal visibility: show a subtle banner for due policies
  useEffect(() => {
    if (!account?.username) { setDuePoliciesCount(0); return; }
    let cancelled = false;
    (async () => {
      try {
        const due = await fetchDuePolicies(account.username);
        if (!cancelled) setDuePoliciesCount(due.length);
      } catch { if (!cancelled) setDuePoliciesCount(0); }
    })();
    return () => { cancelled = true; };
  }, [account?.username]);

  useEffect(() => {
    if (!Array.isArray(batches) || batches.length === 0) return;
    // fetch progress per-batch
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

  // listen for progress update events to refresh progress
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
    window.addEventListener('sunbeth:progressUpdated', h as EventListener);
    return () => {
      window.removeEventListener('sunbeth:progressUpdated', h as EventListener);
    };
  }, [token]);

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
        {duePoliciesCount > 0 && (
          <div className="small" style={{ background: '#f8f9fa', border: '1px solid #e9ecef', padding: 8, borderRadius: 6, marginBottom: 10 }}>
            <span style={{ fontWeight: 600 }}>{duePoliciesCount}</span> policy acknowledgement{duePoliciesCount > 1 ? 's' : ''} due. You may be prompted to review them.
          </div>
        )}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div className="title">Welcome{account ? `, ${account.name}` : ''}</div>
            <div className="muted" style={{ marginTop: 6 }}>You have <strong>{incompleteCount}</strong> pending items</div>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div className="muted small">Due by: <strong style={{ color: 'var(--accent)' }}>{earliestDue ? formatDate(earliestDue) : '—'}</strong></div>
            {rbac.canSeeAdmin && (
              <div style={{ marginTop: 8 }}>
                <Link to="/admin"><button className="btn ghost sm">Admin</button></Link>
              </div>
            )}
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
