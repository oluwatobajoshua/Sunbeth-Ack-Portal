import React, { useState } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { useExternalAuth } from '../context/ExternalAuthContext';

// Unified Login Page for M365 and External Users
const UnifiedLogin: React.FC = () => {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const navigate = useNavigate();
  const location = useLocation();
  const { login: setExternalSession } = useExternalAuth();

  // Handles login for both M365 and external users
  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      // If M365 user, redirect to Microsoft login
      if (email.endsWith('@yourcompany.com')) {
        window.location.href = '/'; // Use your existing M365 login logic
        return;
      }
      // Otherwise, external user: call backend for password auth
      const res = await fetch('/api/external-users/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Login failed');
      // If MFA required, redirect to MFA page
      if (j.mfaRequired) {
        navigate(`/mfa?email=${encodeURIComponent(email)}`);
        return;
      }
  // Otherwise, login successful: store external session
  setExternalSession({ email: j?.email || email, name: j?.name || null });
      navigate('/');
    } catch (e: any) {
      setError(e.message || 'Login failed');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
      <h2>Sign In</h2>
      <form onSubmit={handleLogin}>
        <div style={{ marginBottom: 16 }}>
          <input
            type="email"
            placeholder="Email address"
            value={email}
            onChange={e => setEmail(e.target.value)}
            required
            style={{ width: '100%', padding: 8 }}
            autoFocus
          />
        </div>
        {!email.endsWith('@yourcompany.com') && (
          <div style={{ marginBottom: 16 }}>
            <input
              type="password"
              placeholder="Password"
              value={password}
              onChange={e => setPassword(e.target.value)}
              required
              style={{ width: '100%', padding: 8 }}
            />
          </div>
        )}
        {error && <div style={{ color: 'red', marginBottom: 12 }}>{error}</div>}
        <button className="btn" type="submit" disabled={loading} style={{ width: '100%' }}>
          {loading ? 'Signing in...' : 'Sign In'}
        </button>
      </form>
      <div style={{ marginTop: 16, textAlign: 'center' }}>
        <a href="/forgot-password">Forgot password?</a>
      </div>
    </div>
  );
};

export default UnifiedLogin;
