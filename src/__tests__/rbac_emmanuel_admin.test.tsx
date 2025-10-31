import React from 'react';
import { MemoryRouter } from 'react-router-dom';
import { render, screen, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';

import Layout from '../Layout';
import { AppRoutes } from '../routes';
import { AuthContext } from '../context/AuthContext';
import { RBACContext } from '../context/RBACContext';

// Mock contexts/hooks used by Layout to avoid external dependencies
jest.mock('../context/FeatureFlagsContext', () => ({
  useFeatureFlags: () => ({ externalSupport: false, loaded: true, flags: {}, refresh: jest.fn() })
}));
jest.mock('../context/TenantContext', () => ({
  useTenant: () => ({ tenant: null })
}));
jest.mock('../context/ExternalAuthContext', () => ({
  useExternalAuth: () => ({ user: null, isAuthenticated: false, logout: jest.fn() })
}));

// Ensure runtime config has an API base (RBACProvider will then fetch roles/permissions)
jest.mock('../utils/runtimeConfig', () => ({
  // Core app config used by RBAC and routing
  getApiBase: () => 'https://api.example.test',
  isAdminLight: () => false,
  useAdminModalSelectors: () => true,
  // Branding + busy overlay timings used by Layout's DancingLogoOverlay
  getBrandLogoUrl: () => '',
  getBrandName: () => 'Sunbeth',
  getBrandPrimaryColor: () => '#0a66ff',
  getBusyOverlayShowDelayMs: () => 0,
  getBusyOverlayMinVisibleMs: () => 0
}));

// Mock Graph groups fetch used by RBACProvider
jest.mock('../services/graphService', () => ({
  fetchUserGroups: jest.fn(async () => [])
}));

// Avoid network calls from due policies enforcement during layout mount
jest.mock('../utils/policiesDue', () => ({
  enforceDuePolicies: jest.fn(async () => true)
}));

// Mock backend roles and effective permissions to include Emmanuel as Admin
jest.mock('../services/dbService', () => ({
  getBatches: jest.fn(async () => []),
  getUserProgress: jest.fn(async () => ({ acknowledged: 0, total: 0, percent: 0 })),
  getDocumentsByBatch: jest.fn(async () => []),
  getAcknowledgedDocIds: jest.fn(async () => []),
  getRoles: jest.fn(async () => ([{ email: 'emmanuel.oladayo@sunbeth.net', role: 'Admin' }]))
}));
jest.mock('../services/rbacService', () => ({
  getPermissionCatalog: jest.fn(async () => ([{ key: 'viewAdmin', label: 'View Admin' }])) ,
  getEffectivePermissions: jest.fn(async () => ({ roles: ['Admin'], permissions: { viewAdmin: true, manageSettings: true } }))
}));

// Helper: render app with Auth and real RBACProvider
const renderWithProviders = (account: { name: string; username: string }) => {
  const auth = {
    account,
    token: 'token',
    photo: null,
    login: jest.fn(),
    logout: jest.fn(),
    getToken: jest.fn(async () => 'token')
  } as any;
  const rbac = {
    role: 'Admin',
    canSeeAdmin: true,
    canEditAdmin: true,
    isSuperAdmin: false,
    perms: { viewAdmin: true }
  } as any;
  return render(
    <AuthContext.Provider value={auth}>
      <RBACContext.Provider value={rbac}>
        <MemoryRouter initialEntries={[ '/' ]}>
          <Layout><AppRoutes /></Layout>
        </MemoryRouter>
      </RBACContext.Provider>
    </AuthContext.Provider>
  );
};

describe('RBAC: Emmanuel sees Admin button', () => {
  test('emmanuel.oladayo@sunbeth.net gets Admin View button', async () => {
    renderWithProviders({ name: 'Emmanuel Oladayo', username: 'emmanuel.oladayo@sunbeth.net' });
    await waitFor(async () => {
      const btn = await screen.findByRole('button', { name: /Admin View/i });
      expect(btn).toBeInTheDocument();
    }, { timeout: 3000 });
  });
});
