import React, { useState } from 'react';
import ReactDOM from 'react-dom/client';
import Modal from '../components/Modal';
import PdfViewer from '../components/viewers/PdfViewer';

export type ShowPdfPreviewOptions = {
  title?: string;
  url: string | string[];
  width?: number | string;
  maxWidth?: number | string;
};

/**
 * Programmatically show an in-app PDF preview using our Modal + PdfViewer.
 * Returns a promise that resolves when the user closes the preview.
 */
export function showPdfPreview(opts: ShowPdfPreviewOptions): Promise<void> {
  const { title = 'Legal document', url, width = 920, maxWidth = '95%' } = opts;
  const host = document.createElement('div');
  document.body.appendChild(host);
  const root = ReactDOM.createRoot(host);

  return new Promise<void>((resolve) => {
    const Preview: React.FC = () => {
      const [open, setOpen] = useState(true);
      const firstUrl = Array.isArray(url) ? (url[0] || '') : url;
      // Compute a good download link that hints the server
      const downloadUrl = firstUrl ? (firstUrl + (firstUrl.includes('?') ? '&' : '?') + 'download=1') : '';

      const onClose = () => {
        setOpen(false);
        setTimeout(() => {
          root.unmount();
          host.remove();
          resolve();
        }, 0);
      };

      return (
        <Modal open={open} onClose={onClose} title={title} width={width} maxWidth={maxWidth}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div className="small" style={{ color: '#666' }}>
                You can review the official document here. This is the same file linked in the consent.
              </div>
              <div style={{ flex: 1 }} />
              {firstUrl && (
                <>
                  <a className="btn ghost xs" href={firstUrl} target="_blank" rel="noopener noreferrer">Open in new tab</a>
                  <a className="btn xs" href={downloadUrl} target="_blank" rel="noopener noreferrer">Download</a>
                </>
              )}
            </div>
            <PdfViewer url={url} />
          </div>
        </Modal>
      );
    };

    root.render(
      <React.StrictMode>
        <Preview />
      </React.StrictMode>
    );
  });
}

export default showPdfPreview;
