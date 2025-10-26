// DocumentReader split into presentational parts and hooks; remaining warnings will be addressed incrementally.
/**
 * DocumentReader: Displays a single document within a batch and handles acknowledgement.
 *
 * - Reads documents via dbService (SQLite API or SharePoint Lists).
 * - Sends acknowledgements via flowService.
 * - Navigates previous/next between documents and shows progress.
 */
import React, { useMemo, useState } from 'react';
import { useNavigate, useParams, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
// flow submission and busy indicators are handled inside useAcceptHandler
import Toast from './Toast';
import HeaderBar from './documentReader/HeaderBar';
import ConsentBanner from './documentReader/ConsentBanner';
import GraphAccessHint from './documentReader/GraphAccessHint';
import ViewerFrame from './documentReader/ViewerFrame';
import ActionLinks from './documentReader/ActionLinks';
import AcceptControls from './documentReader/AcceptControls';
import NavControls from './documentReader/NavControls';
import { getApiBase as getApiBaseCfg } from '../utils/runtimeConfig';
import { hasConsent } from '../utils/legalConsent';
import { useBatchAndProgress } from './documentReader/hooks/useBatchAndProgress';
import { useDocUrlResolution } from './documentReader/hooks/useDocUrlResolution';
import { useContentTypeProbe } from './documentReader/hooks/useContentTypeProbe';
import { useAcceptHandler } from './documentReader/hooks/useAcceptHandler';
import { useDocNavigation } from './documentReader/hooks/useDocNavigation';
import { useViewerDecision } from './documentReader/hooks/useViewerDecision';
// progress refresh is handled inside useAcceptHandler

const DocumentReader: React.FC = () => {
  const { id } = useParams();
  const { account, token, getToken } = useAuth();
  const [ack, setAck] = useState(false);
  const title = useMemo(() => `Document ${id}`, [id]);
  const navigate = useNavigate();
  const location = useLocation();
  const params = new URLSearchParams(location.search);
  const batchIdFromQuery = params.get('batchId') || undefined;

  const {
    docs,
    index,
    progressText,
    alreadyAcked,
    ackCheckReady,
    setProgressText,
  } = useBatchAndProgress(id, batchIdFromQuery, token ?? undefined, account?.username || undefined);
  const userName = (account?.name || account?.username || '').toString();
  const [refreshKey, setRefreshKey] = useState<number>(0);

  

  const { onAccept, toastMsg, showToast } = useAcceptHandler({
    ack,
    id,
    title,
    username: account?.username || undefined,
    displayName: account?.name || undefined,
    batchIdFromQuery,
    index,
    docs,
    token,
    navigate,
    setProgressText,
  });

  // loaded via useBatchAndProgress

  const { prevDoc, nextDoc } = useDocNavigation(docs, index, batchIdFromQuery, navigate);

  const currentDoc = (Array.isArray(docs) && index >= 0 && index < docs.length) ? docs[index] : undefined as any;
  const rawUrl = currentDoc?.toba_fileurl || (currentDoc as any)?.url || '';
  // Resolve API base via runtime config; if not set, use same-origin relative '/api'
  const cfgBase = getApiBaseCfg();
  const apiBase = cfgBase ? cfgBase : '';

  const getTokenAdapter = getToken ? (async (scopes?: string[]) => {
    const t = await getToken(scopes);
    return t === null ? undefined : t;
  }) : undefined;
  const { docUrl, backupUrl, needGraphAuth } = useDocUrlResolution(currentDoc, apiBase, getTokenAdapter, refreshKey);
  const contentType = useContentTypeProbe(docUrl, apiBase);

  // resolved via useDocUrlResolution

  // content-type probing via useContentTypeProbe
  // const docTitle = currentDoc?.toba_title || `Document ${id}`;
  // Determine viewer by URL extension or content-type (set by diagnostics)
  const { isPdf, isDocx, proxiedDownloadUrl, openInNewTabUrl, viewerUrls } = useViewerDecision(rawUrl, docUrl, contentType, backupUrl);

  const originalUrl = (currentDoc as any)?.toba_originalurl as string | undefined;

  return (
    <div className="container">
      <div className="card">
        <HeaderBar title={title} />
        <ConsentBanner show={!!(batchIdFromQuery && !hasConsent(account?.username || undefined, batchIdFromQuery))} />
        <GraphAccessHint
          visible={needGraphAuth}
          onGrant={async () => { try { await getToken?.(['Files.Read.All','Sites.Read.All']); setRefreshKey(k => k + 1); } catch (e) { /* noop */ } }}
        />
        <ViewerFrame isPdf={isPdf} isDocx={isDocx} viewerUrls={viewerUrls} docUrl={docUrl} needGraphAuth={needGraphAuth} />
        <ActionLinks docUrl={docUrl} openInNewTabUrl={openInNewTabUrl} proxiedDownloadUrl={proxiedDownloadUrl} originalUrl={originalUrl} />
        <AcceptControls ready={ackCheckReady} alreadyAcked={alreadyAcked} userName={userName} ack={ack} onAckChange={setAck} onAccept={onAccept} />
        <NavControls onPrev={prevDoc} onNext={nextDoc} progressText={progressText} />
        <Toast message={toastMsg} show={showToast} />
      </div>
    </div>
  );
};
export default DocumentReader;
