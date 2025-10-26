import React from 'react';

interface AcceptControlsProps {
  ready: boolean;
  alreadyAcked: boolean;
  userName?: string;
  ack: boolean;
  onAckChange: (checked: boolean) => void;
  onAccept: () => void;
}

const AcceptControls: React.FC<AcceptControlsProps> = ({ ready, alreadyAcked, userName, ack, onAckChange, onAccept }) => {
  if (!ready || alreadyAcked) return null;
  return (
    <div style={{ display: 'flex', gap: 8, marginTop: 12, alignItems: 'center' }}>
      <label className="small">
        <input type="checkbox" checked={ack} onChange={e => onAckChange(e.target.checked)} />{' '}
        {userName
          ? (<><span>I </span><strong>{userName}</strong><span> have read and understood this document.</span></>)
          : 'I have read and understood this document.'}
      </label>
      <div style={{ flex: 1 }} />
      <button className="btn accent sm" id="btnAccept" onClick={onAccept} disabled={!ack}>I Accept</button>
    </div>
  );
};

export default AcceptControls;
