import React, { useState, useEffect, useRef, useMemo } from 'react';
import { Document, Page, pdfjs } from 'react-pdf';
import 'react-pdf/dist/esm/Page/AnnotationLayer.css';
import 'react-pdf/dist/esm/Page/TextLayer.css';

// Set up PDF.js worker
pdfjs.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjs.version}/pdf.worker.min.js`;

interface PdfViewerProps {
  url: string;
}

const PdfViewer: React.FC<PdfViewerProps> = ({ url }) => {
  const [numPages, setNumPages] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const [pdfData, setPdfData] = useState<Uint8Array | null>(null);
  const [page, setPage] = useState<number>(1);
  const [scale, setScale] = useState<number>(1);
  const containerRef = useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = useState<number>(800);
  // IMPORTANT: Use a Blob for pdf.js input to avoid ArrayBuffer detachment when
  // the worker transfers underlying buffers. Keep hook order stable.
  const blobForPdf = useMemo(() => {
    if (!pdfData) return null;
    // Create a fresh copy to ensure an ArrayBuffer (not SharedArrayBuffer) for BlobPart
    const copy = new Uint8Array(pdfData);
    return new Blob([copy.buffer], { type: 'application/pdf' });
  }, [pdfData]);

  useEffect(() => {
    let mounted = true;
    setLoading(true);
    setError('');
    setPdfData(null);

    const fetchWithFallback = async (u: string): Promise<Uint8Array> => {
      const doFetch = async (target: string) => {
        const res = await fetch(target, { headers: { 'Accept': 'application/pdf,*/*;q=0.8' }, cache: 'no-store' as RequestCache });
        if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
        const ab = await res.arrayBuffer();
        return new Uint8Array(ab);
      };

      // 1) Try as-is
      let data = await doFetch(u);
      if (data.byteLength > 0) return data;

      // 2) Retry with download=1 hint (some servers require attachment disposition to stream reliably)
      const dlUrl = u + (u.includes('?') ? '&' : '?') + 'download=1';
      data = await doFetch(dlUrl);
      return data;
    };

    (async () => {
      try {
        const uint8Array = await fetchWithFallback(url);
        if (uint8Array.byteLength === 0) {
          throw new Error('The PDF file is empty, i.e. its size is zero bytes.');
        }
        if (mounted) {
          setPdfData(uint8Array);
          setLoading(false);
        }
      } catch (err) {
        console.error('PDF fetch error:', err);
        if (mounted) {
          setError(err instanceof Error ? err.message : 'Failed to load PDF');
          setLoading(false);
        }
      }
    })();

    return () => {
      mounted = false;
    };
  }, [url]);

  const onDocumentLoadSuccess = ({ numPages }: { numPages: number }) => {
    setNumPages(numPages);
    setPage(1);
  };

  const onDocumentLoadError = (err: Error) => {
    console.error('PDF.js load error:', err);
    setError('Failed to render PDF: ' + err.message);
  };

  // Track container width for fit-to-width and responsive rendering
  useEffect(() => {
    const update = () => {
      const w = containerRef.current?.clientWidth || Math.min(window.innerWidth - 100, 900);
      setContainerWidth(Math.max(320, Math.min(1000, w)));
    };
    update();
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);

  if (loading) {
    return (
      <div style={{ padding: 20, textAlign: 'center', color: '#666' }}>
        Loading PDF...
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ padding: 20, textAlign: 'center', color: '#d32f2f' }}>
        <div style={{ marginBottom: 8 }}>⚠️ Error loading PDF</div>
        <div style={{ fontSize: '0.9em' }}>{error}</div>
      </div>
    );
  }

  if (!pdfData) {
    return (
      <div style={{ padding: 20, textAlign: 'center', color: '#666' }}>
        No PDF data available
      </div>
    );
  }

  const canPrev = page > 1;
  const canNext = page < Math.max(1, numPages);
  const zoomOut = () => setScale(s => Math.max(0.5, Math.round((s - 0.1) * 10) / 10));
  const zoomIn = () => setScale(s => Math.min(3, Math.round((s + 0.1) * 10) / 10));
  const resetZoom = () => setScale(1);
  const goPrev = () => canPrev && setPage(p => p - 1);
  const goNext = () => canNext && setPage(p => p + 1);

  // Width-based sizing for crisp rendering; zoom multiplies the available width
  const pageWidth = Math.floor(containerWidth * scale);

  return (
    <div ref={containerRef} style={{ 
      maxHeight: '70vh', 
      display: 'flex',
      flexDirection: 'column',
      overflow: 'hidden',
      border: '1px solid #ddd', 
      borderRadius: 6,
      background: '#f5f5f5'
    }}>
      {/* Toolbar */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #e6e6e6', background: '#fff' }}>
        <button type="button" className="btn ghost xs" onClick={zoomOut} title="Zoom out">−</button>
        <div className="small" style={{ minWidth: 56, textAlign: 'center' }}>{Math.round(scale * 100)}%</div>
        <button type="button" className="btn ghost xs" onClick={zoomIn} title="Zoom in">+</button>
        <button type="button" className="btn ghost xs" onClick={resetZoom} title="Reset zoom">Reset</button>
        <div style={{ flex: 1 }} />
        <button type="button" className="btn ghost xs" onClick={goPrev} disabled={!canPrev} title="Previous page">←</button>
        <div className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input
            type="number"
            min={1}
            max={Math.max(1, numPages)}
            value={page}
            onChange={(e) => {
              const v = Number(e.target.value || '1');
              if (Number.isFinite(v)) setPage(Math.min(Math.max(1, Math.floor(v)), Math.max(1, numPages)));
            }}
            style={{ width: 64, padding: '2px 6px', border: '1px solid #ddd', borderRadius: 4 }}
          />
          <span>/ {Math.max(1, numPages)}</span>
        </div>
        <button type="button" className="btn ghost xs" onClick={goNext} disabled={!canNext} title="Next page">→</button>
      </div>

      {/* Page viewport, single page at a time for lazy rendering */}
      <div style={{ overflow: 'auto', padding: 12 }}>
        <Document
        file={blobForPdf || undefined}
        onLoadSuccess={onDocumentLoadSuccess}
        onLoadError={onDocumentLoadError}
        loading={
          <div style={{ padding: 20, textAlign: 'center' }}>
            Rendering PDF...
          </div>
        }
      >
          <Page
            key={`page_${page}`}
            pageNumber={page}
            renderTextLayer={true}
            renderAnnotationLayer={true}
            width={pageWidth}
            loading={
              <div style={{ padding: 20, textAlign: 'center', color: '#999' }}>
                Loading page {page}...
              </div>
            }
          />
        </Document>
      </div>
    </div>
  );
};

export default PdfViewer;
