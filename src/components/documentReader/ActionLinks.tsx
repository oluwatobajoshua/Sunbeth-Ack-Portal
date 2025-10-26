import React from 'react';

interface ActionLinksProps {
  docUrl: string;
  openInNewTabUrl: string;
  proxiedDownloadUrl: string;
  originalUrl?: string;
}

const ActionLinks: React.FC<ActionLinksProps> = ({ docUrl, openInNewTabUrl, proxiedDownloadUrl, originalUrl }) => {
  if (!docUrl && !originalUrl) return null;
  return (
    <div className="small" style={{ marginTop: 8, textAlign: 'right' }}>
      {docUrl && (
        <a href={openInNewTabUrl} target="_blank" rel="noopener noreferrer" style={{ marginRight: 12 }}>Open in new tab ↗</a>
      )}
      {docUrl && (
        <a href={proxiedDownloadUrl} className="btn ghost xs">Download</a>
      )}
      {originalUrl && originalUrl !== docUrl && (
        <>
          {docUrl && <span style={{ margin: '0 8px', color: '#bbb' }}>|</span>}
          <a href={originalUrl} target="_blank" rel="noopener noreferrer">View in SharePoint</a>
        </>
      )}
    </div>
  );
};

export default ActionLinks;
