import React, { useState } from 'react';

// MFA Prompt/Setup for External Users
const MFA: React.FC = () => {
  const [email, setEmail] = useState('');
  const [step, setStep] = useState<'setup'|'verify'|'done'>('setup');
  const [secret, setSecret] = useState('');
  const [otpauth, setOtpauth] = useState('');
  const [code, setCode] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  // Parse email from URL if present
  React.useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    setEmail(params.get('email') || '');
  }, []);

  // Start MFA setup (get secret and otpauth)
  const startSetup = async () => {
    setError(null);
    setLoading(true);
    try {
      const res = await fetch('/api/external-users/mfa/setup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Failed to start MFA setup');
      setSecret(j.secret);
      setOtpauth(j.otpauth);
      setStep('verify');
    } catch (e: any) {
      setError(e.message || 'Failed to start MFA setup');
    } finally {
      setLoading(false);
    }
  };

  // Verify MFA code
  const handleVerify = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);
    try {
      const res = await fetch('/api/external-users/mfa/verify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, code })
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || 'Failed to verify code');
      setStep('done');
    } catch (e: any) {
      setError(e.message || 'Failed to verify code');
    } finally {
      setLoading(false);
    }
  };

  if (step === 'done') {
    return (
      <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
        <h2>MFA Enabled</h2>
        <div>MFA is now enabled for your account. You may now log in.</div>
        <a href="/login"><button className="btn" style={{ marginTop: 16 }}>Go to Login</button></a>
      </div>
    );
  }

  if (step === 'setup') {
    return (
      <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
        <h2>Enable MFA</h2>
        <div style={{ marginBottom: 16 }}>
          <button className="btn" onClick={startSetup} disabled={loading || !email}>
            {loading ? 'Starting...' : 'Start MFA Setup'}
          </button>
        </div>
        {error && <div style={{ color: 'red', marginBottom: 12 }}>{error}</div>}
      </div>
    );
  }

  if (step === 'verify') {
    return (
      <div className="container" style={{ maxWidth: 400, margin: '0 auto', padding: 32 }}>
        <h2>Verify MFA</h2>
        <div style={{ marginBottom: 16 }}>
          <div>Scan this QR code in your authenticator app:</div>
          <img src={`https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(otpauth)}&size=200x200`} alt="MFA QR" style={{ margin: '16px 0' }} />
          <div style={{ fontSize: 12, wordBreak: 'break-all' }}>Or enter secret: <b>{secret}</b></div>
        </div>
        <form onSubmit={handleVerify}>
          <div style={{ marginBottom: 16 }}>
            <input
              type="text"
              placeholder="Enter 6-digit code"
              value={code}
              onChange={e => setCode(e.target.value)}
              required
              pattern="\\d{6}"
              style={{ width: '100%', padding: 8 }}
              autoFocus
            />
          </div>
          {error && <div style={{ color: 'red', marginBottom: 12 }}>{error}</div>}
          <button className="btn" type="submit" disabled={loading} style={{ width: '100%' }}>
            {loading ? 'Verifying...' : 'Verify & Enable MFA'}
          </button>
        </form>
      </div>
    );
  }

  return null;
};

export default MFA;
