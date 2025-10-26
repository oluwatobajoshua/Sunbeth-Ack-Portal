/* eslint-disable max-lines-per-function, complexity */
import { useEffect, useState } from 'react';

export function useContentTypeProbe(docUrl: string, apiBase: string) {
  const [contentType, setContentType] = useState<string>('');

  useEffect(() => {
    (async () => {
      try {
        setContentType('');
        if (!docUrl) return;
        const isAbsoluteApi = apiBase && typeof docUrl === 'string' && docUrl.startsWith(apiBase);
        const isRelativeApi = typeof docUrl === 'string' && docUrl.startsWith('/api/');
        if (isAbsoluteApi || isRelativeApi) {
          const abs = isAbsoluteApi ? docUrl : ((apiBase || '') + docUrl);
          const diagUrl = abs + (abs.includes('?') ? '&' : '?') + 'diag=1';
          const res = await fetch(diagUrl, { cache: 'no-store' });
          const info = await res.json().catch(() => null as any);
          const ct = (info?.mime || info?.contentType || res.headers.get('content-type') || '').toString();
          if (ct) setContentType(ct);
        }
      } catch (e) {
        // ignore; viewer selection will fall back to extension-based detection
      }
    })();
  }, [docUrl, apiBase]);

  return contentType;
}
