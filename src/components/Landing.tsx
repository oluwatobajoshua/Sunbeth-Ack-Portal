import React, { useEffect, useState } from 'react';
import { Link } from 'react-router-dom';
import { useAuth } from '../context/AuthContext';
import { useFeatureFlags } from '../context/FeatureFlagsContext';

const Landing: React.FC = () => {
  const { login } = useAuth();
  const navigate = require('react-router-dom').useNavigate();
  const { externalSupport, loaded } = useFeatureFlags();
  const [signingIn, setSigningIn] = useState(false);
  const [variant, setVariant] = useState<string>(() => { try { return localStorage.getItem('landing_variant') || 'regular'; } catch { return 'regular'; } });
  useEffect(() => {
    const read = () => { try { setVariant(localStorage.getItem('landing_variant') || 'regular'); } catch { /* noop */ } };
    const onStorage = (e: StorageEvent) => { if (!e.key || e.key === 'landing_variant') read(); };
    window.addEventListener('storage', onStorage as EventListener);
    return () => {
      window.removeEventListener('storage', onStorage as EventListener);
    };
  }, []);
  return (
    <div className="container" style={{ display: 'flex', justifyContent: 'center' }}>
      <div className={`card hero ${variant === 'compact' ? 'compact' : ''}`} style={{ maxWidth: 920, width: '100%' }}>
        <div className="hero-grid" style={{ alignItems: 'center' }}>
          <div className="hero-text" style={{ textAlign: 'left' }}>
            <div className="eyebrow">Welcome to</div>
            <h1 className="hero-title" style={{ fontSize: 32, lineHeight: 1.2 }}>Sunbeth Document Acknowledgement</h1>
            <p className="hero-lead" style={{ marginTop: 12 }}>Read, understand, and acknowledge the policies and updates that keep our people and business safe.</p>

            <ul className="features" style={{ marginTop: 14 }}>
              <li><strong>Clear workflows</strong> — Step-by-step acknowledgement and progress tracking.</li>
              <li><strong>Secure</strong> — Single sign-on with your corporate account.</li>
              <li><strong>Compliant</strong> — Keep a verifiable record of acknowledgements.</li>
            </ul>

            <div style={{ marginTop: 22, display: 'flex', gap: 10 }}>
              <button
                className="btn sm"
                onClick={() => {
                  // If external support is disabled, skip the gateway and start SSO immediately
                  if (loaded && !externalSupport) {
                    if (!signingIn) { setSigningIn(true); }
                    login();
                  } else {
                    navigate('/login');
                  }
                }}
                disabled={loaded && !externalSupport && signingIn}
                aria-busy={loaded && !externalSupport && signingIn}
              >
                {loaded && !externalSupport && signingIn ? 'Opening Microsoft sign-in…' : 'Sign in to get started'}
              </button>
              <Link to="/about"><button className="btn ghost sm" type="button">Learn more</button></Link>
            </div>
          </div>

          <div className="hero-visual" style={{ display: 'flex', justifyContent: 'center' }}>
            <img src="/images/landing_image.png" alt="Sunbeth Document Acknowledgement" style={{ width: '100%', maxWidth: 420, borderRadius: 10, boxShadow: '0 4px 24px rgba(0,0,0,0.10)' }} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default Landing;
