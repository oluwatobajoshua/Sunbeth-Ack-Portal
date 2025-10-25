import React from 'react';
import { BrowserRouter } from 'react-router-dom';
import { MsalProvider } from '@azure/msal-react';
import { msalInstance } from './services/msalConfig';
import { AuthProvider } from './context/AuthContext';
import { ExternalAuthProvider } from './context/ExternalAuthContext';
import { RBACProvider } from './context/RBACContext';
import { FeatureFlagsProvider } from './context/FeatureFlagsContext';
import { TenantProvider } from './context/TenantContext';
import { AppRoutes, RouteChangeLogger } from './routes';
import Layout from './Layout';
import ThemeController from './context/ThemeController';

const App: React.FC = () => {
  const tree = (
    <TenantProvider>
      <ExternalAuthProvider>
        <AuthProvider>
          <RBACProvider>
            <FeatureFlagsProvider>
              <BrowserRouter future={{ v7_startTransition: true, v7_relativeSplatPath: true } as any}>
                <RouteChangeLogger />
                <ThemeController />
                <Layout>
                  <AppRoutes />
                </Layout>
              </BrowserRouter>
            </FeatureFlagsProvider>
          </RBACProvider>
        </AuthProvider>
      </ExternalAuthProvider>
    </TenantProvider>
  );

  // Always provide MSAL context
  return (
    <MsalProvider instance={msalInstance}>
      {tree}
    </MsalProvider>
  );
};

export default App;
