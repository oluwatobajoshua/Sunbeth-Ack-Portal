import React from 'react';

interface GraphAccessHintProps {
  visible: boolean;
  onGrant: () => void;
}

const GraphAccessHint: React.FC<GraphAccessHintProps> = ({ visible, onGrant }) => {
  if (!visible) return null;
  return (
    <div className="small" style={{ marginBottom: 8, background: '#fff8e1', border: '1px solid #ffe0b2', padding: 10, borderRadius: 8 }}>
      This document is stored in SharePoint. We need Microsoft Graph access to preview it here.
      <button className="btn ghost xs" style={{ marginLeft: 8 }} onClick={onGrant}>
        Grant access
      </button>
    </div>
  );
};

export default GraphAccessHint;
