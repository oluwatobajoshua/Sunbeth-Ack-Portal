import React, { useState } from 'react';

// Password Reset for External Users
const ResetPassword: React.FC = () => {
  const [email, setEmail] = useState('');
  const [token, setToken] = useState('');
  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [step, setStep] = useState<'request'|'reset'|'done'>('request');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  // Parse email/token from URL if present
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get('email')) setEmail(params.get('email')!);
    if (params.get('token')) {
      setToken(params.get('token')!);
      setStep('reset');
    }
  }, []);

  const handleRequest = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      const res = await fetch('/api/external-users/request-reset', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Failed to send reset email');
      setStep('reset');
    } catch (e: any) {
      setError(e.message || 'Failed to send reset email');
    } finally {
      setLoading(false);
    }
  };

  const handleReset = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (password !== confirm) {
      setError('Passwords do not match');
      return;
    }
    setLoading(true);
    try {
      const res = await fetch('/api/external-users/reset-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, token, password })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Failed to reset password');
      setStep('done');
    } catch (e: any) {
      setError(e.message || 'Failed to reset password');
    } finally {
      setLoading(false);
    }
  };

  if (step === 'done') {
    return (
      <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
        <h2>Password Reset</h2>
        <div>Your password has been reset. You may now log in.</div>
        <a href="/login"><button className="btn" style={{ marginTop: 16 }}>Go to Login</button></a>
      </div>
    );
  }

  return (
    <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
      <h2>{step === 'request' ? 'Reset Password' : 'Set New Password'}</h2>
      {step === 'request' ? (
        <form onSubmit={handleRequest}>
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
          {error && <div style={{ color: 'red', marginBottom: 12 }}>{error}</div>}
          <button className="btn" type="submit" disabled={loading} style={{ width: '100%' }}>
            {loading ? 'Sending...' : 'Send Reset Email'}
          </button>
        </form>
      ) : (
        <form onSubmit={handleReset}>
          <input type="hidden" value={email} />
          <input type="hidden" value={token} />
          <div style={{ marginBottom: 16 }}>
            <input
              type="password"
              placeholder="New password"
              value={password}
              onChange={e => setPassword(e.target.value)}
              required
              minLength={8}
              style={{ width: '100%', padding: 8 }}
            />
          </div>
          <div style={{ marginBottom: 16 }}>
            <input
              type="password"
              placeholder="Confirm password"
              value={confirm}
              onChange={e => setConfirm(e.target.value)}
              required
              minLength={8}
              style={{ width: '100%', padding: 8 }}
            />
          </div>
          {error && <div style={{ color: 'red', marginBottom: 12 }}>{error}</div>}
          <button className="btn" type="submit" disabled={loading} style={{ width: '100%' }}>
            {loading ? 'Resetting...' : 'Reset Password'}
          </button>
        </form>
      )}
    </div>
  );
};

export default ResetPassword;
