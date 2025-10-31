import React from 'react';
import { render, screen } from '@testing-library/react';
import '@testing-library/jest-dom';

import AdminPanel from '../components/AdminPanel';
import { AuthContext } from '../context/AuthContext';
import { RBACContext } from '../context/RBACContext';

// Keep external flags and tenants quiet in tests
jest.mock('../context/FeatureFlagsContext', () => ({
  useFeatureFlags: () => ({ externalSupport: false, loaded: true, flags: {}, refresh: jest.fn() })
}));

// Disable SQLite-dependent behaviors to avoid network
jest.mock('../utils/runtimeConfig', () => ({
  isSQLiteEnabled: () => false,
  getApiBase: () => null,
  isAdminLight: () => true,
  useAdminModalSelectors: () => true,
  // Overlay/branding (spares any components that might use them)
  getBrandLogoUrl: () => '',
  getBrandName: () => 'Sunbeth',
  getBrandPrimaryColor: () => '#0a66ff',
  getBusyOverlayShowDelayMs: () => 0,
  getBusyOverlayMinVisibleMs: () => 0
}));

// Minimal auth provider value
const auth = {
  account: { name: 'Test User', username: 'test.user@sunbeth.net' },
  token: 'token',
  photo: null,
  login: jest.fn(),
  logout: jest.fn(),
  getToken: jest.fn(async () => 'token')
} as any;

const renderWithRBAC = (rbac: any) => {
  return render(
    <AuthContext.Provider value={auth}>
      <RBACContext.Provider value={rbac}>
        <AdminPanel />
      </RBACContext.Provider>
    </AuthContext.Provider>
  );
};

describe('Permissions Matrix gating', () => {
  test('viewAdmin permission grants access even when canSeeAdmin=false', async () => {
    const rbac = {
      role: 'Employee',
      canSeeAdmin: false,
      canEditAdmin: false,
      isSuperAdmin: false,
      perms: { viewAdmin: true, editBatch: true }
    };
    renderWithRBAC(rbac);
    expect(screen.queryByText(/Access Denied/i)).not.toBeInTheDocument();
    expect(screen.getByText(/Admin Panel/i)).toBeInTheDocument();
  // "Create Batch" tab should appear with editBatch permission
  expect(screen.getByRole('tab', { name: /Create Batch/i })).toBeInTheDocument();
  });

  test('no viewAdmin and no role yields Access Denied', async () => {
    const rbac = {
      role: 'Employee',
      canSeeAdmin: false,
      canEditAdmin: false,
      isSuperAdmin: false,
      perms: {}
    };
    renderWithRBAC(rbac);
    expect(screen.getByText(/Access Denied/i)).toBeInTheDocument();
  });
});
