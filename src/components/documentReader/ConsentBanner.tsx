import React from 'react';

interface ConsentBannerProps {
  show: boolean;
}

const ConsentBanner: React.FC<ConsentBannerProps> = ({ show }) => {
  if (!show) return null;
  return (
    <div className="small" style={{ marginTop: 8, background: '#fff8e1', border: '1px solid #ffe0b2', padding: 10, borderRadius: 8 }}>
      To submit an acknowledgement in this batch, you must first consent to legal terms. You&apos;ll be prompted when you click I Accept.
    </div>
  );
};

export default ConsentBanner;
