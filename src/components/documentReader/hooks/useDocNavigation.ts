import { useCallback } from 'react';

export function useDocNavigation(
  docs: Array<{ toba_documentid: string }> ,
  index: number,
  batchIdFromQuery: string | undefined,
  navigate: (to: string) => void,
) {
  const prevDoc = useCallback(() => {
    if (!Array.isArray(docs) || index <= 0) {
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const prevId = docs[index - 1].toba_documentid;
    navigate(`/document/${prevId}?batchId=${batchIdFromQuery}`);
  }, [docs, index, batchIdFromQuery, navigate]);

  const nextDoc = useCallback(() => {
    if (!Array.isArray(docs) || index >= docs.length - 1) {
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const nextId = docs[index + 1].toba_documentid;
    navigate(`/document/${nextId}?batchId=${batchIdFromQuery}`);
  }, [docs, index, batchIdFromQuery, navigate]);

  return { prevDoc, nextDoc };
}
