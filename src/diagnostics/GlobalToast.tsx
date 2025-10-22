import React, { useEffect, useState } from 'react';
import Toast from '../components/Toast';

export const GlobalToast: React.FC = () => {
  const [msg, setMsg] = useState('');
  const [show, setShow] = useState(false);

  useEffect(() => {
    const handler = (e: any) => {
      const m = e?.detail?.message || e?.detail || String(e);
      setMsg(m);
      setShow(true);
      window.setTimeout(() => setShow(false), 2000);
    };
    window.addEventListener('sunbeth:toast', handler as EventListener);
    return () => window.removeEventListener('sunbeth:toast', handler as EventListener);
  }, []);

  return <Toast message={msg} show={show} />;
};

export default GlobalToast;
