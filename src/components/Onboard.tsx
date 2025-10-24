import React, { useState } from 'react';

// Onboarding: Set Password after invite
const Onboard: React.FC = () => {
  const [email, setEmail] = useState('');
  const [token, setToken] = useState('');
  const [password, setPassword] = useState('');
  const [confirm, setConfirm] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [loading, setLoading] = useState(false);

  // Parse email/token from URL if present
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    setEmail(params.get('email') || '');
    setToken(params.get('token') || '');
  }, []);

  const handleSetPassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (password !== confirm) {
      setError('Passwords do not match');
      return;
    }
    setLoading(true);
    try {
      const res = await fetch('/api/external-users/set-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, token, password })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Failed to set password');
      setSuccess(true);
    } catch (e: any) {
      setError(e.message || 'Failed to set password');
    } finally {
      setLoading(false);
    }
  };

  if (success) {
    return (
      <div className="wrap centered">
        <div className="grid" style={{ maxWidth: 680, margin: '0 auto' }}>
          <main>
            <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
              <div style={{ background: 'linear-gradient(135deg, var(--primary) 0%, var(--primary-500, #0a6b55) 100%)', color: '#fff', padding: '18px 22px' }}>
                <div style={{ fontWeight: 800, fontSize: 18 }}>Password Set</div>
                <div className="small" style={{ opacity: .9 }}>You may now log in</div>
              </div>
              <div style={{ padding: 18 }}>
                <div>Your password has been set. You may now log in.</div>
                <a href="/login"><button className="btn" style={{ marginTop: 16 }}>Go to Login</button></a>
              </div>
            </div>
          </main>
        </div>
      </div>
    );
  }

  return (
    <div className="wrap centered">
      <div className="grid" style={{ maxWidth: 680, margin: '0 auto' }}>
        <main>
          <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
            <div style={{ background: 'linear-gradient(135deg, var(--primary) 0%, var(--primary-500, #0a6b55) 100%)', color: '#fff', padding: '18px 22px' }}>
              <div style={{ fontWeight: 800, fontSize: 18 }}>Set Your Password</div>
              <div className="small" style={{ opacity: .9 }}>Complete your onboarding to access Sunbeth</div>
            </div>
            <div style={{ padding: 18 }}>
              <form onSubmit={handleSetPassword}>
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
                    className="form-control"
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
                    className="form-control"
                  />
                </div>
                {error && <div className="small" style={{ color: '#b91c1c', background: '#fef2f2', border: '1px solid #fee2e2', padding: 10, borderRadius: 8, marginBottom: 12 }}>{error}</div>}
                <button className="btn" type="submit" disabled={loading} style={{ width: '100%' }}>
                  {loading ? 'Setting password...' : 'Set Password'}
                </button>
              </form>
              <div style={{ marginTop: 10 }}>
                <a href="/login">Back to login</a>
              </div>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default Onboard;
