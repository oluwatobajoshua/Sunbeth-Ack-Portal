import React, { useState, useEffect, useRef } from 'react';
import { renderAsync } from 'docx-preview';

interface DocxViewerProps {
  url: string;
}

const DocxViewer: React.FC<DocxViewerProps> = ({ url }) => {
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>('');
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    let mounted = true;
    setLoading(true);
    setError('');

    (async () => {
      try {
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const blob = await response.blob();
        
        if (mounted && containerRef.current) {
          // Clear previous content
          containerRef.current.innerHTML = '';
          
          await renderAsync(blob, containerRef.current, undefined, {
            className: 'docx-preview',
            inWrapper: false,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            ignoreLastRenderedPageBreak: true,
            experimental: false,
            trimXmlDeclaration: true,
            useBase64URL: false,
            renderHeaders: true,
            renderFooters: true,
            renderFootnotes: true,
            renderEndnotes: true
          });
          
          setLoading(false);
        }
      } catch (err) {
        console.error('DOCX fetch/render error:', err);
        if (mounted) {
          setError(err instanceof Error ? err.message : 'Failed to load DOCX');
          setLoading(false);
        }
      }
    })();

    return () => {
      mounted = false;
    };
  }, [url]);

  if (loading) {
    return (
      <div style={{ padding: 20, textAlign: 'center', color: '#666' }}>
        Loading document...
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ padding: 20, textAlign: 'center', color: '#d32f2f' }}>
        <div style={{ marginBottom: 8 }}>⚠️ Error loading document</div>
        <div style={{ fontSize: '0.9em' }}>{error}</div>
      </div>
    );
  }

  return (
    <div style={{ 
      maxHeight: '70vh', 
      overflowY: 'auto', 
      border: '1px solid #ddd', 
      borderRadius: 6,
      padding: 12,
      background: '#fff'
    }}>
      <div ref={containerRef} style={{ padding: 8 }} />
    </div>
  );
};

export default DocxViewer;
