/* eslint-disable max-lines-per-function, complexity, max-depth */
import { useState, useCallback } from 'react';
import { busyPush, busyPop } from '../../../utils/busy';
import { sendAcknowledgement } from '../../../services/flowService';
import { getUserProgress } from '../../../services/dbService';
import { requestConsentIfNeeded } from '../../../utils/legalConsent';

interface AcceptDeps {
  ack: boolean;
  id?: string;
  title: string;
  username?: string;
  displayName?: string;
  batchIdFromQuery?: string;
  index: number;
  docs: Array<{ toba_documentid: string }>;
  token?: string | null;
  navigate: (to: string) => void;
  setProgressText: (s: string) => void;
}

export function useAcceptHandler({
  ack,
  id,
  title,
  username,
  displayName,
  batchIdFromQuery,
  index,
  docs,
  token,
  navigate,
  setProgressText,
}: AcceptDeps) {
  const [toastMsg, setToastMsg] = useState('');
  const [showToast, setShowToast] = useState(false);

  const onAccept = useCallback(async () => {
    if (!ack) return;
    try {
      const ok = await requestConsentIfNeeded(username || undefined, batchIdFromQuery || undefined);
      if (!ok) return;
    } catch (e) { return; }

    const payload = {
      userDisplay: displayName,
      userEmail: username,
      userPrincipalName: username,
      email: username,
      documentId: id,
      documentTitle: title,
      batchId: batchIdFromQuery || '1',
      batchName: 'unknown',
      ackmethod: 'Clicked Accept',
    };

    try {
      busyPush('Submitting your acknowledgement...');
      setToastMsg('Acknowledgement submitted');
      setShowToast(true);
      await sendAcknowledgement(payload);
      try {
        const p = await getUserProgress(payload.batchId, token ?? undefined, undefined, username || undefined);
        setProgressText(`${p.percent}%`);
        if (p.percent >= 100) {
          navigate(`/summary?batchId=${payload.batchId}`);
        } else {
          if (index < docs.length - 1) {
            const nextId = docs[index + 1].toba_documentid;
            navigate(`/document/${nextId}?batchId=${payload.batchId}`);
          } else {
            navigate(`/batch/${payload.batchId}`);
          }
        }
      } catch (e) { /* noop */ }
      setTimeout(() => setShowToast(false), 1200);
    } catch (e) {
      navigate(`/batch/${payload.batchId}`);
    } finally {
      busyPop();
    }
  }, [ack, username, batchIdFromQuery, displayName, id, title, index, docs, token, navigate, setProgressText]);

  return { onAccept, toastMsg, showToast };
}
