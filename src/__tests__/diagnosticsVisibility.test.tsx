import React from 'react';
import { render, screen } from '@testing-library/react';
import '@testing-library/jest-dom';

import AdminPanel from '../components/AdminPanel';
import { AuthContext } from '../context/AuthContext';
import { RBACContext } from '../context/RBACContext';

// Mock feature flags to avoid network calls in provider
jest.mock('../context/FeatureFlagsContext', () => ({
  useFeatureFlags: () => ({ externalSupport: false, flags: {}, loaded: true, refresh: jest.fn() })
}));

// Ensure runtime config signals SQLite mode with a stable API base
jest.mock('../utils/runtimeConfig', () => ({
  isSQLiteEnabled: () => true,
  getApiBase: () => 'https://sunbeth-ack-portal-backend.vercel.app',
  isAdminLight: () => false,
  useAdminModalSelectors: () => true
}));

// Basic fetch mock for endpoints touched on mount
beforeAll(() => {
  const mockFetch = jest.fn(async (input: RequestInfo | URL) => {
    const url = typeof input === 'string' ? input : input.toString();
    if (url.endsWith('/api/health')) {
      return { ok: true, json: async () => ({ ok: true }) } as any;
    }
    if (url.endsWith('/api/stats')) {
      return { ok: true, json: async () => ({ totalBatches: 0, activeBatches: 0, totalUsers: 0, completionRate: 0, overdueBatches: 0, avgCompletionTime: 0 }) } as any;
    }
    // default harmless response
    return { ok: true, json: async () => ({}), text: async () => '' } as any;
  });
  (globalThis as any).fetch = mockFetch;
  if (typeof window !== 'undefined') { (window as any).fetch = mockFetch; }
});

afterAll(() => {
  const f = (globalThis as any).fetch;
  if (f && typeof f.mockClear === 'function') f.mockClear();
});

const commonAuth = {
  account: { name: 'User', username: 'user@sunbeth.net' },
  token: 'token',
  photo: null,
  login: jest.fn(),
  logout: jest.fn(),
  getToken: jest.fn(async () => 'token')
} as any;

const renderWithRBAC = (rbac: any) => {
  return render(
    <AuthContext.Provider value={commonAuth}>
      <RBACContext.Provider value={rbac}>
        <AdminPanel />
      </RBACContext.Provider>
    </AuthContext.Provider>
  );
};

describe('Diagnostics tab visibility', () => {
  test('is hidden for non-super-admin users', async () => {
    renderWithRBAC({
      role: 'Admin',
      canSeeAdmin: true,
      canEditAdmin: true,
      isSuperAdmin: false,
      perms: { exportAnalytics: true }
    });
  const maybeDiag = screen.queryByRole('tab', { name: /system diagnostics/i });
    expect(maybeDiag).not.toBeInTheDocument();
  });

  test('is visible for super admin users', async () => {
    renderWithRBAC({
      role: 'Admin',
      canSeeAdmin: true,
      canEditAdmin: true,
      isSuperAdmin: true,
      perms: { exportAnalytics: true }
    });
  const diagTab = await screen.findByRole('tab', { name: /system diagnostics/i });
    expect(diagTab).toBeInTheDocument();
  });
});
