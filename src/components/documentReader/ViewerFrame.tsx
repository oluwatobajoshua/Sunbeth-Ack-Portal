import React from 'react';
import PdfViewer from '../viewers/PdfViewer';
import DocxViewer from '../viewers/DocxViewer';
import GenericInline from '../viewers/GenericInline';

interface ViewerFrameProps {
  isPdf: boolean;
  isDocx: boolean;
  viewerUrls: string | string[];
  docUrl: string;
  needGraphAuth: boolean;
}

const ViewerFrame: React.FC<ViewerFrameProps> = ({ isPdf, isDocx, viewerUrls, docUrl, needGraphAuth }) => {
  return (
    <div className="viewer" style={{ marginTop: 12 }}>
      {docUrl ? (
        isPdf ? (
          <PdfViewer url={viewerUrls} />
        ) : isDocx ? (
          <DocxViewer url={Array.isArray(viewerUrls) ? viewerUrls[0] : viewerUrls} />
        ) : (
          // Try a generic inline viewer via iframe; backend proxy clears X-Frame-Options
          <GenericInline url={viewerUrls} />
        )
      ) : (
        <div className="muted small" style={{ padding: 12, textAlign: 'center' }}>
          {needGraphAuth ? 'Please grant access to preview this SharePoint document.' : 'No document URL found for this item.'}
        </div>
      )}
    </div>
  );
};

export default ViewerFrame;
