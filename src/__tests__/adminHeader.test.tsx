import React from 'react';
import { render, screen, waitFor, within } from '@testing-library/react';
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

const renderAdmin = () => {
  const auth = {
    account: { name: 'Admin User', username: 'admin@sunbeth.net' },
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
    isSuperAdmin: true,
    perms: { exportAnalytics: true }
  } as any;

  return render(
    <AuthContext.Provider value={auth}>
      <RBACContext.Provider value={rbac}>
        <AdminPanel />
      </RBACContext.Provider>
    </AuthContext.Provider>
  );
};

describe('AdminPanel header', () => {
  test('shows Backend API indicator with health link', async () => {
    renderAdmin();

    // Host from mocked getApiBase
    const hostText = 'sunbeth-ack-portal-backend.vercel.app';
    const link = await screen.findByRole('link', { name: hostText });
    expect(link).toBeInTheDocument();
    expect(link).toHaveAttribute('href', 'https://sunbeth-ack-portal-backend.vercel.app/api/health');

    // Copy button present
    const copyBtn = screen.getByRole('button', { name: /copy/i });
    expect(copyBtn).toBeInTheDocument();

    // Trigger a manual refresh to ensure our mocked fetch runs in this test turn
  const apiHealthBox = screen.getByTitle('SQLite API health');
  const refresh = within(apiHealthBox).getByRole('button', { name: /refresh/i });
  refresh.click();

    // API health badge should be present and we should have pinged /api/health
    expect(within(apiHealthBox).getByText(/API:/i)).toBeInTheDocument();
    await waitFor(() => {
      const mf = (globalThis as any).fetch as jest.Mock;
      const calledHealth = mf.mock.calls.some(([arg]) => {
        if (typeof arg === 'string') return /\/api\/health$/.test(arg);
        try {
          const url = (arg && (arg as any).url) ? String((arg as any).url) : String(arg);
          return /\/api\/health$/.test(url);
        } catch { return false; }
      });
      expect(calledHealth).toBe(true);
    }, { timeout: 3000 });
  });
});
