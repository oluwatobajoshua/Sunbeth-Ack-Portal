/* eslint-disable max-lines-per-function, complexity */
import { useEffect, useState } from 'react';
import { getDocumentsByBatch, getUserProgress, getAcknowledgedDocIds, getDocumentById } from '../../../services/dbService';
import type { Doc } from '../../../types/models';

export interface BatchState {
  docs: Doc[];
  index: number;
  progressText: string;
  alreadyAcked: boolean;
  ackCheckReady: boolean;
}

export interface UseBatchAndProgressResult extends BatchState {
  setIndex: (i: number) => void;
  setProgressText: (t: string) => void;
  setAlreadyAcked: (v: boolean) => void;
  setAckCheckReady: (v: boolean) => void;
}

export function useBatchAndProgress(
  id: string | undefined,
  batchIdFromQuery: string | undefined,
  token: string | undefined | null,
  username: string | undefined,
): UseBatchAndProgressResult {
  const [docs, setDocs] = useState<Doc[]>([]);
  const [index, setIndex] = useState<number>(0);
  const [progressText, setProgressText] = useState<string>('—');
  const [alreadyAcked, setAlreadyAcked] = useState<boolean>(false);
  const [ackCheckReady, setAckCheckReady] = useState<boolean>(false);

  useEffect(() => {
    (async () => {
      try {
        if (batchIdFromQuery) {
          const batchId = batchIdFromQuery as string;
          const list = await getDocumentsByBatch(batchId);
          setDocs(list);
          const idx = list.findIndex((d: Doc) => d.toba_documentid === id);
          setIndex(idx >= 0 ? idx : 0);
          try {
            const p = await getUserProgress(batchId, token ?? undefined, undefined, username);
            setProgressText(`${p.percent}%`);
          } catch (e) { /* noop */ }
          try {
            const ackIds = await getAcknowledgedDocIds(batchId, token ?? undefined, username);
            setAlreadyAcked(id ? ackIds.includes(id) : false);
          } catch (e) { /* noop */ }
          finally { setAckCheckReady(true); }
        } else if (id) {
          const doc = await getDocumentById(id);
          if (doc) {
            const mapped: Doc = {
              toba_documentid: String((doc as any).toba_documentid || (doc as any).id || id),
              toba_title: (doc as any).toba_title || (doc as any).title || `Document ${id}`,
              toba_version: String((doc as any).toba_version || (doc as any).version || '1'),
              toba_requiressignature: !!(((doc as any).toba_requiressignature) ?? (doc as any).requiresSignature ?? false),
              toba_fileurl: (doc as any).toba_fileurl || (doc as any).url,
            } as any;
            setDocs([mapped]);
            setIndex(0);
            setAckCheckReady(true);
          } else {
            setAckCheckReady(true);
          }
        } else {
          setAckCheckReady(true);
        }
      } catch (e) {
        setAckCheckReady(true);
      }
    })();
  }, [batchIdFromQuery, id, token, username]);

  return {
    docs,
    index,
    progressText,
    alreadyAcked,
    ackCheckReady,
    setIndex,
    setProgressText,
    setAlreadyAcked,
    setAckCheckReady,
  };
}
