import React from 'react';
import { BrowserRouter } from 'react-router-dom';
import { MsalProvider } from '@azure/msal-react';
import { msalInstance } from './services/msalConfig';
import { AuthProvider } from './context/AuthContext';
import { RBACProvider } from './context/RBACContext';
import { AppRoutes, RouteChangeLogger } from './routes';
import Layout from './Layout';
import { useRuntimeMock } from './utils/runtimeMock';

const App: React.FC = () => {
  const runtimeMock = useRuntimeMock();
  const tree = (
    <AuthProvider>
      <RBACProvider>
        <BrowserRouter future={{ v7_startTransition: true, v7_relativeSplatPath: true } as any}>
          <RouteChangeLogger />
          <Layout>
            <AppRoutes />
          </Layout>
        </BrowserRouter>
      </RBACProvider>
    </AuthProvider>
  );

  // Only provide MSAL context in live mode; in mock mode msal-react will not be used
  return runtimeMock ? tree : (
    <MsalProvider instance={msalInstance}>
      {tree}
    </MsalProvider>
  );
};

export default App;
