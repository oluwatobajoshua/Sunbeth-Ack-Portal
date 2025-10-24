import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useFeatureFlags } from '../context/FeatureFlagsContext';

const LoginGateway: React.FC = () => {
  const { login } = useAuth();
  const navigate = useNavigate();
  const { externalSupport, loaded } = useFeatureFlags();
  const [autoSigningIn, setAutoSigningIn] = React.useState(false);

  // If external support is disabled, skip the gateway and trigger SSO immediately
  React.useEffect(() => {
    if (loaded && !externalSupport && !autoSigningIn) {
      setAutoSigningIn(true);
      // Trigger SSO immediately
      Promise.resolve().then(() => login());
    }
  }, [loaded, externalSupport, login, autoSigningIn]);

  return (
    <div className="wrap centered">
      <div className="grid" style={{ maxWidth: 980, margin: '0 auto' }}>
        <main>
          <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
            <div style={{
              background: 'linear-gradient(135deg, var(--primary) 0%, var(--primary-500, #0a6b55) 100%)',
              color: '#fff',
              padding: '20px 22px'
            }}>
              <div style={{ fontWeight: 800, fontSize: 18 }}>Sign in to Sunbeth</div>
              <div className="small" style={{ opacity: .9 }}>Choose your account type</div>
            </div>

            <div style={{ padding: 18 }}>
              {autoSigningIn && (
                <div className="small" style={{
                  marginBottom: 12,
                  background: '#f5fbf9',
                  border: '1px solid #d1f0e6',
                  color: '#0a6b55',
                  padding: 8,
                  borderRadius: 8
                }}>
                  Opening Microsoft sign-in…
                </div>
              )}
              <div style={{ display: 'grid', gap: 14 }}>
                <div className="card" style={{ boxShadow: 'none', border: '1px solid #eef1ee' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
                    <div>
                      <div className="title" style={{ margin: 0 }}>Internal Employee</div>
                      <div className="small" style={{ color: 'var(--muted)' }}>Sign in with Microsoft 365 (SSO)</div>
                    </div>
                    <button className="btn sm" onClick={() => login()} disabled={autoSigningIn} aria-busy={autoSigningIn}>
                      {autoSigningIn ? 'Opening…' : 'Sign in with Microsoft'}
                    </button>
                  </div>
                </div>

                {externalSupport && (
                  <div className="card" style={{ boxShadow: 'none', border: '1px solid #eef1ee' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12 }}>
                      <div>
                        <div className="title" style={{ margin: 0 }}>External User</div>
                        <div className="small" style={{ color: 'var(--muted)' }}>Sign in with email and password</div>
                      </div>
                      <button className="btn accent sm" onClick={() => navigate('/login/external')}>Continue as External</button>
                    </div>
                  </div>
                )}

                <div className="small" style={{ color: 'var(--muted-2)' }}>
                  Need help? Contact IT Support.
                </div>
              </div>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
};

export default LoginGateway;
