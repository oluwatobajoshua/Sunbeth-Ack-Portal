/**
 * DocumentReader: Displays a single document within a batch and handles acknowledgement.
 *
 * - Reads documents via dbService (SQLite API or SharePoint Lists).
 * - Sends acknowledgements via flowService.
 * - Navigates previous/next between documents and shows progress.
 */
import React, { useMemo, useState, useEffect } from 'react';
import PdfViewer from './viewers/PdfViewer';
import DocxViewer from './viewers/DocxViewer';
import { Link, useNavigate, useParams, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { sendAcknowledgement } from '../services/flowService';
import { busyPush, busyPop } from '../utils/busy';
import Toast from './Toast';
import { getDocumentsByBatch, getUserProgress, getAcknowledgedDocIds, getDocumentById } from '../services/dbService';
import type { Doc } from '../types/models';
import { getApiBase as getApiBaseCfg } from '../utils/runtimeConfig';

const DocumentReader: React.FC = () => {
  const { id } = useParams();
  const { account, token, getToken } = useAuth();
  const [ack, setAck] = useState(false);
  const title = useMemo(() => `Document ${id}`, [id]);
  const [docs, setDocs] = useState<Doc[]>([]);
  const [index, setIndex] = useState<number>(0);
  const [progressText, setProgressText] = useState<string>('—');
  const [alreadyAcked, setAlreadyAcked] = useState<boolean>(false);
  const [ackCheckReady, setAckCheckReady] = useState<boolean>(false);
  const userName = (account?.name || account?.username || '').toString();
  const [needGraphAuth, setNeedGraphAuth] = useState<boolean>(false);
  const [refreshKey, setRefreshKey] = useState<number>(0);

  const navigate = useNavigate();
  const location = useLocation();
  const params = new URLSearchParams(location.search);
  const batchIdFromQuery = params.get('batchId') || undefined;

  const onAccept = async () => {
    if (!ack) return;
    const payload = { 
      userDisplay: account?.name,
      userEmail: account?.username,
      userPrincipalName: account?.username,
      email: account?.username,
      documentId: id,
      documentTitle: title,
      batchId: batchIdFromQuery || '1',
      batchName: 'unknown',
      ackmethod: 'Clicked Accept'
    };
    try {
      busyPush('Submitting your acknowledgement...');
      setToastMsg('Acknowledgement submitted');
      setShowToast(true);
      await sendAcknowledgement(payload);
      // ensure progress updated
      try {
        const p = await getUserProgress(payload.batchId, token ?? undefined, undefined, account?.username || undefined);
        setProgressText(`${p.percent}%`);
        // Navigate only if batch is completed; otherwise advance to next or return to batch
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
      } catch {}
      setTimeout(() => setShowToast(false), 1200);
    } catch (e) {
      // fallback behavior
      navigate(`/batch/${payload.batchId}`);
    }
    finally {
      busyPop();
    }
  };

  const [toastMsg, setToastMsg] = React.useState('');
  const [showToast, setShowToast] = React.useState(false);

  // load docs for the batch and find current index; if batchId is missing, fall back to resolving by document id
  useEffect(() => {
    (async () => {
      try {
        if (batchIdFromQuery) {
          const batchId = batchIdFromQuery as string;
          const list = await getDocumentsByBatch(batchId, token ?? undefined);
          setDocs(list);
          const idx = list.findIndex((d: Doc) => d.toba_documentid === id);
          setIndex(idx >= 0 ? idx : 0);
          // Progress and ack state
          try {
            const p = await getUserProgress(batchId, token ?? undefined, undefined, account?.username || undefined);
            setProgressText(`${p.percent}%`);
          } catch {}
          try {
            const ackIds = await getAcknowledgedDocIds(batchId, token ?? undefined, account?.username || undefined);
            setAlreadyAcked(ackIds.includes(id!));
          } catch {}
          finally { setAckCheckReady(true); }
        } else if (id) {
          const doc = await getDocumentById(id);
          if (doc) {
            const mapped: Doc = {
              toba_documentid: String(doc.toba_documentid || doc.id || id),
              toba_title: doc.toba_title || doc.title || `Document ${id}`,
              toba_version: String(doc.toba_version || doc.version || '1'),
              toba_requiressignature: !!(doc.toba_requiressignature ?? doc.requiresSignature ?? false),
              toba_fileurl: doc.toba_fileurl || doc.url,
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
  }, [batchIdFromQuery, id, token, account?.username]);

  const prevDoc = () => {
    if (!Array.isArray(docs) || index <= 0) {
      // go back to batch
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const prevId = docs[index - 1].toba_documentid;
    navigate(`/document/${prevId}?batchId=${batchIdFromQuery}`);
  };

  const nextDoc = () => {
    if (!Array.isArray(docs) || index >= docs.length - 1) {
      // last doc -> go back to batch; summary is shown only when batch is fully acknowledged
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const nextId = docs[index + 1].toba_documentid;
    navigate(`/document/${nextId}?batchId=${batchIdFromQuery}`);
  };

  const currentDoc = (Array.isArray(docs) && index >= 0 && index < docs.length) ? docs[index] : undefined as any;
  const rawUrl = currentDoc?.toba_fileurl || (currentDoc as any)?.url || '';
  // Resolve API base via runtime config; if not set, use same-origin relative '/api'
  const cfgBase = getApiBaseCfg();
  const apiBase = cfgBase ? cfgBase : '';

  const [docUrl, setDocUrl] = useState<string>(rawUrl);
  const [contentType, setContentType] = useState<string>('');

  useEffect(() => {
    (async () => {
      try {
        if (!currentDoc) { setDocUrl(''); return; }
        // Prefer Graph proxy for SharePoint-sourced docs when IDs are available
        const driveId = (currentDoc as any)?.toba_driveid || (currentDoc as any)?.driveId;
        const itemId = (currentDoc as any)?.toba_itemid || (currentDoc as any)?.itemId;
        const source = (currentDoc as any)?.toba_source || (currentDoc as any)?.source;
        setNeedGraphAuth(false);
        if (apiBase && source === 'sharepoint' && driveId && itemId) {
          try {
            const at = await getToken?.(['Files.Read.All', 'Sites.Read.All']);
            if (at) {
              setDocUrl(`${apiBase}/api/proxy/graph?driveId=${encodeURIComponent(driveId)}&itemId=${encodeURIComponent(itemId)}&token=${encodeURIComponent(at)}`);
              return;
            }
          } catch {
            // Token not available; prompt user to grant Graph access
            setNeedGraphAuth(true);
            setDocUrl('');
            return;
          }
        }
        // If looks like a SharePoint URL but IDs are missing, try Graph shares API by URL
        const url = currentDoc?.toba_fileurl || (currentDoc as any)?.url || '';
        const looksSharePoint = /\.sharepoint\.com\//i.test(url) || /graph\.microsoft\.com\//i.test(url);
        if (apiBase && looksSharePoint) {
          try {
            const at = await getToken?.(['Files.Read.All', 'Sites.Read.All']);
            if (at) {
              setDocUrl(`${apiBase}/api/proxy/graph?url=${encodeURIComponent(url)}&token=${encodeURIComponent(at)}`);
              return;
            }
          } catch {
            // Token not available; do not fall back to unauthenticated proxy (would fail). Show hint.
            setNeedGraphAuth(true);
            setDocUrl('');
            return;
          }
        }
        // Fallback: simple URL proxy
        setDocUrl(apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(url)}` : url);
      } catch {
        const url = currentDoc?.toba_fileurl || (currentDoc as any)?.url || '';
        setDocUrl(apiBase ? `${apiBase}/api/proxy?url=${encodeURIComponent(url)}` : url);
      }
    })();
  // re-evaluate when the selected doc changes or apiBase changes
  }, [currentDoc, apiBase, refreshKey]);

  // Attempt a quick proxy diagnostics call to learn the content-type so we can pick the correct viewer
  useEffect(() => {
    (async () => {
      try {
        setContentType('');
        if (!docUrl) return;
        if (apiBase && typeof docUrl === 'string' && docUrl.startsWith(apiBase)) {
          const diagUrl = docUrl + (docUrl.includes('?') ? '&' : '?') + 'diag=1';
          const res = await fetch(diagUrl);
          const info = await res.json().catch(() => null);
          const ct = (info?.contentType || res.headers.get('content-type') || '').toString();
          if (ct) setContentType(ct);
        }
      } catch {
        // ignore; viewer selection will fall back to extension-based detection
      }
    })();
  }, [docUrl, apiBase]);
  const docTitle = currentDoc?.toba_title || `Document ${id}`;
  // Determine viewer by URL extension or content-type (set by diagnostics)
  const extHintPdf = /\.pdf(\?|#|$)/i.test(rawUrl) || /\.pdf(\?|#|$)/i.test(docUrl);
  const extHintDocx = /\.docx(\?|#|$)/i.test(rawUrl) || /\.docx(\?|#|$)/i.test(docUrl);
  const isPdf = extHintPdf || /application\/pdf/i.test(contentType);
  const isDocx = extHintDocx || /vnd\.openxmlformats-officedocument\.wordprocessingml\.document/i.test(contentType);
  const proxiedDownloadUrl = docUrl ? (docUrl + (docUrl.includes('?') ? '&' : '?') + 'download=1') : '';

  return (
    <div className="container">
      <div className="card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div className="title">{title}</div>
            <div className="muted small">Please read and acknowledge</div>
          </div>
          <Link to="/"><button className="btn ghost sm">← Back</button></Link>
        </div>
        <div className="viewer" style={{ marginTop: 12 }}>
          {needGraphAuth && (
            <div className="small" style={{ marginBottom: 8, background: '#fff8e1', border: '1px solid #ffe0b2', padding: 10, borderRadius: 8 }}>
              This document is stored in SharePoint. We need Microsoft Graph access to preview it here.
              <button className="btn ghost xs" style={{ marginLeft: 8 }}
                onClick={async () => { try { await getToken?.(['Files.Read.All','Sites.Read.All']); setRefreshKey(k => k + 1); } catch {} }}>
                Grant access
              </button>
            </div>
          )}
          {docUrl ? (
            isPdf ? (
              <PdfViewer url={docUrl} />
            ) : isDocx ? (
              <DocxViewer url={docUrl} />
            ) : (
              <div className="muted small" style={{ padding: 12, textAlign: 'center', border: '1px solid #eee', borderRadius: 6 }}>
                Preview not available for this file type. Use Download or Open in new tab.
              </div>
            )
          ) : (
            <div className="muted small" style={{ padding: 12, textAlign: 'center' }}>
              {needGraphAuth ? 'Please grant access to preview this SharePoint document.' : 'No document URL found for this item.'}
            </div>
          )}
          {(docUrl || rawUrl) && (
            <div className="small" style={{ marginTop: 8, textAlign: 'right' }}>
              {docUrl && (
                <a href={docUrl} target="_blank" rel="noopener noreferrer" style={{ marginRight: 12 }}>Open in new tab ↗</a>
              )}
              {docUrl && (
                <a href={proxiedDownloadUrl} className="btn ghost xs">Download</a>
              )}
            </div>
          )}
        </div>

        {/* Render the accept controls only when acknowledgment check has completed and doc is not already acknowledged */}
        {ackCheckReady && !alreadyAcked && (
          <div style={{ display: 'flex', gap: 8, marginTop: 12, alignItems: 'center' }}>
            <label className="small">
              <input type="checkbox" onChange={e => setAck(e.target.checked)} />{' '}
              {userName
                ? (<><span>I </span><strong>{userName}</strong><span> have read and understood this document.</span></>)
                : 'I have read and understood this document.'}
            </label>
            <div style={{ flex: 1 }} />
            <button className="btn accent sm" id="btnAccept" onClick={onAccept} disabled={!ack}>I Accept</button>
          </div>
        )}

        <div style={{ marginTop: 12 }}>
          <div className="controls">
            <button className="btn ghost sm" id="btnPrev" onClick={prevDoc}>← Previous</button>
            <div className="spacer" />
            <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <button className="btn ghost sm" id="btnNext" onClick={nextDoc}>Next →</button>
            </div>
          </div>

          <div className="progressBar" aria-hidden="true"><i style={{ width: progressText }} /></div>
          <div className="muted small" style={{ marginTop: 8 }}>{progressText} complete</div>
        </div>
      </div>
    </div>
  );
};
export default DocumentReader;
