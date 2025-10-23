import React, { useEffect, useState } from 'react';
import { Link, useLocation } from 'react-router-dom';
import { sendAcknowledgement } from '../services/flowService';
import { alertSuccess, alertError } from '../utils/alerts';
import { getUserProgress } from '../services/dbService';
import { useAuth } from '../context/AuthContext';

const Summary: React.FC = () => {
  const { token, account } = useAuth();
  const loc = useLocation();
  const qs = new URLSearchParams(loc.search);
  const batchId = qs.get('batchId') || undefined;
  const [percent, setPercent] = useState<number | null>(null);
  const [nudgeStatus, setNudgeStatus] = useState<'idle'|'sending'|'sent'|'error'>('idle');
  const handleNudge = async () => {
    if (!batchId || !account) return;
    setNudgeStatus('sending');
    try {
      // Fetch all documents for this batch
      const docs = await import('../services/dbService').then(m => m.getDocumentsByBatch(batchId));
      if (!Array.isArray(docs) || docs.length === 0) throw new Error('No documents found for batch');
      // Send acknowledgement for each document
      for (const doc of docs) {
        await sendAcknowledgement({
          batchId,
          documentId: doc.toba_documentid || doc.id || doc.documentId,
          userDisplay: account.name,
          userEmail: account.username,
          userPrincipalName: account.username,
          email: account.username,
          ackmethod: 'Nudge Admin (manual)'
        });
      }
      setNudgeStatus('sent');
      await alertSuccess('Notification Sent', 'The admin has been notified that you have completed your batch.');
    } catch (err) {
      setNudgeStatus('error');
      await alertError('Notification Failed', 'There was a problem sending the notification. Please try again.');
    }
  };

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
        {isComplete && (
          <div style={{ marginBottom: 12 }}>
            <button className="btn sm" onClick={handleNudge} disabled={nudgeStatus==='sending' || nudgeStatus==='sent'}>
              {nudgeStatus === 'idle' && 'Notify Admin'}
              {nudgeStatus === 'sending' && 'Sending...'}
              {nudgeStatus === 'sent' && 'Notification Sent!'}
              {nudgeStatus === 'error' && 'Error, Try Again'}
            </button>
          </div>
        )}
        <Link to="/"><button className="btn">Return to Dashboard</button></Link>
      </div>
    </div>
  );
};
export default Summary;
