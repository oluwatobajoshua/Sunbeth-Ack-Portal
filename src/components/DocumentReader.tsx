/**
 * DocumentReader: Displays a single document within a batch and handles acknowledgement.
 *
 * - Reads documents via dbService (SQLite API or SharePoint Lists).
 * - Sends acknowledgements via flowService.
 * - Navigates previous/next between documents and shows progress.
 */
import React, { useMemo, useState, useEffect } from 'react';
import { Link, useNavigate, useParams, useLocation } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { sendAcknowledgement } from '../services/flowService';
import Toast from './Toast';
import { getDocumentsByBatch, getUserProgress, getAcknowledgedDocIds } from '../services/dbService';
import type { Doc } from '../types/models';

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
  };

  const [toastMsg, setToastMsg] = React.useState('');
  const [showToast, setShowToast] = React.useState(false);

  // load docs for the batch and find current index
  useEffect(() => {
    (async () => {
      if (!batchIdFromQuery) return;
      try {
        const batchId = batchIdFromQuery as string;
        const list = await getDocumentsByBatch(batchId, token ?? undefined);
        setDocs(list);
        const idx = list.findIndex((d: Doc) => d.toba_documentid === id);
        setIndex(idx >= 0 ? idx : 0);
        // set a progress text
        try {
          const p = await getUserProgress(batchId, token ?? undefined, undefined, account?.username || undefined);
          setProgressText(`${p.percent}%`);
        } catch {}
        // check if this doc is already acknowledged
        try {
          const ackIds = await getAcknowledgedDocIds(batchId, token ?? undefined, account?.username || undefined);
          setAlreadyAcked(ackIds.includes(id!));
        } catch {}
        finally { setAckCheckReady(true); }
      } catch (e) {
        // ignore
        setAckCheckReady(true);
      }
    })();
  }, [batchIdFromQuery, id, token, account?.username]);

  const prevDoc = () => {
    if (index <= 0) {
      // go back to batch
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const prevId = docs[index - 1].toba_documentid;
    navigate(`/document/${prevId}?batchId=${batchIdFromQuery}`);
  };

  const nextDoc = () => {
    if (index >= docs.length - 1) {
      // last doc -> go back to batch; summary is shown only when batch is fully acknowledged
      navigate(`/batch/${batchIdFromQuery || ''}`);
      return;
    }
    const nextId = docs[index + 1].toba_documentid;
    navigate(`/document/${nextId}?batchId=${batchIdFromQuery}`);
  };

  const currentDoc = docs[index];
  const rawUrl = currentDoc?.toba_fileurl || (currentDoc as any)?.url || '';
  const rawBase = process.env.REACT_APP_API_BASE || '';
  // Resolve API base robustly to avoid requests hitting the frontend origin by mistake
  let apiBase = '' as string;
  if (/^https?:\/\//i.test(rawBase)) {
    apiBase = rawBase.replace(/\/$/, '');
  } else {
    // Try a configured base in localStorage (optional manual override)
    const stored = (() => { try { return localStorage.getItem('sunbeth:apiBase') || ''; } catch { return ''; }})();
    if (stored && /^https?:\/\//i.test(stored)) {
      apiBase = stored.replace(/\/$/, '');
    } else {
      // Default fallback for local dev even if not on localhost hostname
      apiBase = 'http://localhost:4000';
    }
  }

  const [docUrl, setDocUrl] = useState<string>(rawUrl);

  useEffect(() => {
    (async () => {
      try {
        if (!currentDoc) { setDocUrl(''); return; }
        // Prefer Graph proxy for SharePoint-sourced docs when IDs are available
        const driveId = (currentDoc as any)?.toba_driveid || (currentDoc as any)?.driveId;
        const itemId = (currentDoc as any)?.toba_itemid || (currentDoc as any)?.itemId;
        const source = (currentDoc as any)?.toba_source || (currentDoc as any)?.source;
        if (apiBase && source === 'sharepoint' && driveId && itemId) {
          try {
            const at = await getToken?.(['Files.Read.All', 'Sites.Read.All']);
            if (at) {
              setDocUrl(`${apiBase}/api/proxy/graph?driveId=${encodeURIComponent(driveId)}&itemId=${encodeURIComponent(itemId)}&token=${encodeURIComponent(at)}`);
              return;
            }
          } catch {
            // fall through to URL proxy
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
            // ignore and fall back
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
  }, [currentDoc, apiBase]);
  const docTitle = currentDoc?.toba_title || `Document ${id}`;
  const isPdf = typeof docUrl === 'string' && /\.pdf(\?|#|$)/i.test(docUrl);

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
          {docUrl ? (
            <iframe
              key={docUrl}
              title={docTitle}
              src={docUrl}
              style={{ width: '100%', height: '70vh', border: '1px solid #eee', borderRadius: 6 }}
              sandbox="allow-same-origin allow-scripts allow-forms allow-popups"
            />
          ) : (
            <div className="muted small" style={{ padding: 12, textAlign: 'center' }}>
              No document URL found for this item.
            </div>
          )}
          {rawUrl && (
            <div className="small" style={{ marginTop: 8, textAlign: 'right' }}>
              <a href={rawUrl} target="_blank" rel="noopener noreferrer">Open in new tab ↗</a>
            </div>
          )}
        </div>

        {/* Render the accept controls only when acknowledgment check has completed and doc is not already acknowledged */}
        {ackCheckReady && !alreadyAcked && (
          <div style={{ display: 'flex', gap: 8, marginTop: 12, alignItems: 'center' }}>
            <label className="small"><input type="checkbox" onChange={e => setAck(e.target.checked)} /> I have read and understood this document.</label>
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
