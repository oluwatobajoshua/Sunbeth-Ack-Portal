import React from 'react';
import { BrowserRouter } from 'react-router-dom';
import { MsalProvider } from '@azure/msal-react';
import { msalInstance } from './services/msalConfig';
import { AuthProvider } from './context/AuthContext';
import { RBACProvider } from './context/RBACContext';
import { FeatureFlagsProvider } from './context/FeatureFlagsContext';
import { AppRoutes, RouteChangeLogger } from './routes';
import Layout from './Layout';

const App: React.FC = () => {
  const tree = (
    <AuthProvider>
      <RBACProvider>
        <FeatureFlagsProvider>
          <BrowserRouter future={{ v7_startTransition: true, v7_relativeSplatPath: true } as any}>
            <RouteChangeLogger />
            <Layout>
              <AppRoutes />
            </Layout>
          </BrowserRouter>
        </FeatureFlagsProvider>
      </RBACProvider>
    </AuthProvider>
  );

  // Always provide MSAL context
  return (
    <MsalProvider instance={msalInstance}>
      {tree}
    </MsalProvider>
  );
};

export default App;
