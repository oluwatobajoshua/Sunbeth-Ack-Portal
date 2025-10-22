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

export const AppRoutes: React.FC = () => {
  const { account } = useAuth();
  return (
    <Routes>
      <Route path="/" element={account ? <Dashboard /> : <Landing />} />
      <Route path="/about" element={<About />} />
      <Route path="/batch/:id" element={<RequireAuth><BatchDetail /></RequireAuth>} />
      <Route path="/batch/:id/completed" element={<RequireAuth><CompletedBatch /></RequireAuth>} />
      <Route path="/document/:id" element={<RequireAuth><DocumentReader /></RequireAuth>} />
      <Route path="/summary" element={<RequireAuth><Summary /></RequireAuth>} />
      <Route path="/admin" element={<AdminGuard><AdminPanel /></AdminGuard>} />
    </Routes>
  );
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
