import React from 'react';
import { Routes, Route, useLocation, Navigate } from 'react-router-dom';
import { info } from './diagnostics/logger';
import { useRBAC } from './context/RBACContext';
import Dashboard from './components/Dashboard';
import BatchDetail from './components/BatchDetail';
import DocumentReader from './components/DocumentReader';
import Summary from './components/Summary';
import CompletedBatch from './components/CompletedBatch';
import AdminPanel from './components/AdminPanel';
import Landing from './components/Landing';
import About from './components/About';
import { useAuth } from './context/AuthContext';
import { useFeatureFlags } from './context/FeatureFlagsContext';

import LoginGateway from './components/LoginGateway.tsx';
import ExternalLogin from './components/ExternalLogin.tsx';
import Logout from './components/Logout.tsx';

export const AppRoutes: React.FC = () => {
  const { account } = useAuth();
  const { externalSupport, loaded } = useFeatureFlags();
  const UnifiedLogin = React.lazy(() => import('./components/UnifiedLogin'));
  const Onboard = React.lazy(() => import('./components/Onboard'));
    const ResetPassword = React.lazy(() => import('./components/ResetPassword'));
    const MFA = React.lazy(() => import('./components/MFA'));
  const ExternalEnabledGuard: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    if (!loaded) return null;
    if (!externalSupport) return <Navigate to="/" replace />;
    return <>{children}</>;
  };

  return (
    <Routes>
      <Route path="/" element={account ? <Dashboard /> : <Landing />} />
      <Route path="/login" element={<React.Suspense fallback={null}><LoginGateway /></React.Suspense>} />
      <Route path="/login/external" element={<React.Suspense fallback={null}><ExternalEnabledGuard><ExternalLogin /></ExternalEnabledGuard></React.Suspense>} />
      {/* Onboarding and password setup for external users */}
      <Route path="/onboard" element={<React.Suspense fallback={null}><ExternalEnabledGuard><Onboard /></ExternalEnabledGuard></React.Suspense>} />
      <Route path="/mfa" element={<React.Suspense fallback={null}><ExternalEnabledGuard><MFA /></ExternalEnabledGuard></React.Suspense>} />
      <Route path="/reset-password" element={<React.Suspense fallback={null}><ExternalEnabledGuard><ResetPassword /></ExternalEnabledGuard></React.Suspense>} />
      <Route path="/about" element={<About />} />
      <Route path="/batch/:id" element={<RequireAuth><BatchDetail /></RequireAuth>} />
      <Route path="/batch/:id/completed" element={<RequireAuth><CompletedBatch /></RequireAuth>} />
      <Route path="/document/:id/*" element={<RequireAuth><DocumentReader /></RequireAuth>} />
      <Route path="/summary" element={<RequireAuth><Summary /></RequireAuth>} />
      <Route path="/admin" element={<AdminGuard><AdminPanel /></AdminGuard>} />
    </Routes>
  );
      <Route path="/logout" element={<Logout />} />
};

// small component to log route changes
export const RouteChangeLogger: React.FC = () => {
  const loc = useLocation();
  React.useEffect(() => { info('Route changed', { pathname: loc.pathname }); }, [loc.pathname]);
  return null;
};
const AdminGuard: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const rbac = useRBAC();
  if (!rbac.canSeeAdmin) return <Navigate to="/" replace />;
  return <>{children}</>;
};

const RequireAuth: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { account } = useAuth();
  if (!account) return <Navigate to="/" replace />;
  return <>{children}</>;
};
