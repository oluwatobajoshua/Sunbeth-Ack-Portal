import React from 'react';

interface GenericInlineProps {
  url: string | string[];
  height?: number;
}

// Simple iframe-based inline viewer for browser-supported types (images, text, html, audio/video players)
const GenericInline: React.FC<GenericInlineProps> = ({ url, height = 600 }) => {
  const src = Array.isArray(url) ? (url[0] || '') : url;
  if (!src) return null;
  return (
    <div style={{ border: '1px solid #eee', borderRadius: 6, overflow: 'hidden' }}>
      <iframe
        title="Document Preview"
        src={src}
        style={{ width: '100%', height }}
        sandbox="allow-same-origin allow-scripts allow-popups allow-downloads"
      />
    </div>
  );
};

export default GenericInline;
