import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';

const Logout: React.FC = () => {
  const { account, logout } = useAuth();
  const navigate = useNavigate();

  React.useEffect(() => {
    (async () => {
      try {
        if (account) {
          await logout(); // Internal (MSAL) sign-out
        }
      } finally {
        navigate('/', { replace: true }); // External or post-internal -> landing
      }
    })();
  }, [account, logout, navigate]);

  return (
    <div className="container" style={{ maxWidth: 420, margin: '0 auto', padding: 24 }}>
      <div className="card" style={{ padding: 16 }}>
        <div className="title">Signing you out…</div>
        <div className="muted small">Please wait.</div>
      </div>
    </div>
  );
};

export default Logout;
