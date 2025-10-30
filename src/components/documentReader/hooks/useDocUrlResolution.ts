/* eslint-disable max-lines-per-function, complexity, max-depth */
import { useEffect, useState } from 'react';

export function useDocUrlResolution(
  currentDoc: any,
  apiBase: string,
  getToken?: (scopes?: string[]) => Promise<string | undefined>,
  refreshKey?: number,
) {
  const [docUrl, setDocUrl] = useState<string>('');
  const [backupUrl, setBackupUrl] = useState<string>('');
  const [needGraphAuth, setNeedGraphAuth] = useState<boolean>(false);

  useEffect(() => {
    (async () => {
      try {
        if (!currentDoc) { setDocUrl(''); return; }
        const driveId = (currentDoc as any)?.toba_driveid || (currentDoc as any)?.driveId;
        const itemId = (currentDoc as any)?.toba_itemid || (currentDoc as any)?.itemId;
        const source = (currentDoc as any)?.toba_source || (currentDoc as any)?.source;
        setNeedGraphAuth(false);
        const url0 = (currentDoc as any)?.toba_fileurl || (currentDoc as any)?.url || '';
        // Treat placeholders like /api/files/undefined as invalid
        const looksLocal = typeof url0 === 'string' && ((apiBase && url0.startsWith(apiBase + '/api/files/')) || url0.startsWith('/api/files/'));
  const validLocal = looksLocal && (/\/api\/files\/[0-9]+(?:[/?#]|$)/).test(String(url0).replace(String(apiBase||''), ''));
        const isLocalFile = !!validLocal;
        if (isLocalFile) {
          setDocUrl(url0);
          const orig = (currentDoc as any)?.toba_originalurl || (currentDoc as any)?.url || '';
          setBackupUrl(orig ? (apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(orig)}` : orig) : '');
          return;
        }
        if (apiBase && source === 'sharepoint' && driveId && itemId) {
          try {
            const at = await getToken?.(['Files.Read.All', 'Sites.Read.All']);
            if (at) {
              setDocUrl(`${apiBase}/api/proxy/graph?driveId=${encodeURIComponent(driveId)}&itemId=${encodeURIComponent(itemId)}&token=${encodeURIComponent(at)}`);
              const orig = (currentDoc as any)?.toba_originalurl || (currentDoc as any)?.url || '';
              setBackupUrl(orig ? (apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(orig)}` : orig) : '');
              return;
            }
          } catch (e) {
            setNeedGraphAuth(true);
            setDocUrl('');
            setBackupUrl('');
            return;
          }
        }
        const url = (currentDoc as any)?.toba_fileurl || (currentDoc as any)?.url || '';
        const looksSharePoint = /\.sharepoint\.com\//i.test(url) || /graph\.microsoft\.com\//i.test(url);
        if (apiBase && looksSharePoint) {
          try {
            const at = await getToken?.(['Files.Read.All', 'Sites.Read.All']);
            if (at) {
              setDocUrl(`${apiBase}/api/proxy/graph?url=${encodeURIComponent(url)}&token=${encodeURIComponent(at)}`);
              setBackupUrl(apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(url)}` : url);
              return;
            }
          } catch (e) {
            setNeedGraphAuth(true);
            setDocUrl('');
            setBackupUrl('');
            return;
          }
        }
        setDocUrl(apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(url)}` : url);
        setBackupUrl(url);
      } catch (e) {
        const url = (currentDoc as any)?.toba_fileurl || (currentDoc as any)?.url || '';
        setDocUrl(apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(url)}` : url);
        setBackupUrl(url);
      }
    })();
  }, [currentDoc, apiBase, refreshKey, getToken]);

  return { docUrl, backupUrl, needGraphAuth, setNeedGraphAuth };
}
