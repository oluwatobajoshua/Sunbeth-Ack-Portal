import React from 'react';

interface KPIStatProps {
  label: string;
  value: React.ReactNode;
  color?: string;
}

const KPIStat: React.FC<KPIStatProps> = ({ label, value, color }) => {
  return (
    <div className="card" style={{ padding: 16, textAlign: 'center' }}>
      <div style={{ fontSize: 28, fontWeight: 'bold', color: color || 'var(--primary)' }}>{value}</div>
      <div className="small muted">{label}</div>
    </div>
  );
};

export default KPIStat;
