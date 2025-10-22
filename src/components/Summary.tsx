import React, { useEffect, useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import { getUserProgress } from '../services/dbService';
import { useAuth } from '../context/AuthContext';

const Summary: React.FC = () => {
  const { token, account } = useAuth();
  const loc = useLocation();
  const qs = new URLSearchParams(loc.search);
  const batchId = qs.get('batchId') || undefined;
  const [percent, setPercent] = useState<number | null>(null);

  useEffect(() => {
    let active = true;
    (async () => {
      if (!batchId) return; // no specific batch context
      try {
        const p = await getUserProgress(batchId, token ?? undefined, undefined, account?.username || undefined);
        if (active) setPercent(p.percent);
      } catch {
        if (active) setPercent(null);
      }
    })();
    return () => { active = false; };
  }, [batchId, token, account?.username]);

  const isComplete = percent !== null && percent >= 100;

  return (
    <div className="container">
      <div className="card" style={{ textAlign: 'center' }}>
        <div className="title">{isComplete ? '✅ Batch Completed' : 'In Progress'}</div>
        <div style={{ fontWeight: 700, color: 'var(--primary)', marginTop: 8 }}>{batchId || '—'}</div>
        <div className="small muted" style={{ marginTop: 8 }}>
          {isComplete ? 'All documents acknowledged.' : (percent === null ? 'Checking progress…' : `${percent}% acknowledged. Keep going!`)}
        </div>
        <div style={{ height: 14 }} />
        <Link to="/"><button className="btn">Return to Dashboard</button></Link>
      </div>
    </div>
  );
};
export default Summary;
