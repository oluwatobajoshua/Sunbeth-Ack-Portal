import React from 'react';
import { MemoryRouter, Route, Routes } from 'react-router-dom';
import { render, screen, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';

import Layout from '../Layout';
import { AppRoutes } from '../routes';
import { AuthContext } from '../context/AuthContext';
import { RBACContext } from '../context/RBACContext';

// Mock db service methods used by components under test
jest.mock('../services/dbService', () => ({
  getBatches: jest.fn(async () => []),
  getDocumentsByBatch: jest.fn(async () => [{ toba_documentid: 'd1', toba_title: 'Doc 1' }]),
  getUserProgress: jest.fn(async () => ({ acknowledged: 0, total: 1, percent: 0 })),
  getAcknowledgedDocIds: jest.fn(async (batchId: string) => batchId === '1' ? ['d1'] : [])
}));

// Helper: render with minimal providers
const renderWithProviders = (ui: React.ReactNode, { account = null }: { account?: any } = {}) => {
  const auth = {
    account,
    token: account ? 'token' : null,
    photo: null,
    login: jest.fn(),
    logout: jest.fn(),
    getToken: jest.fn(async () => (account ? 'token' : null))
  };
  const rbac = { role: 'Employee', canSeeAdmin: false, canEditAdmin: false };
  return render(
    <AuthContext.Provider value={auth as any}>
      <RBACContext.Provider value={rbac as any}>
        {ui}
      </RBACContext.Provider>
    </AuthContext.Provider>
  );
};

describe('UI/UX and navigation', () => {
  beforeAll(() => {
    // Ensure mock mode behaviors for components relying on env
    (process as any).env = { ...(process as any).env, REACT_APP_USE_MOCK: 'true' };
  });

  test('Landing visible when unauthenticated at root', async () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/']}>
        <Layout><AppRoutes /></Layout>
      </MemoryRouter>
    );
    expect(await screen.findByText(/Sign in to get started/i)).toBeInTheDocument();
  });

  test('Unauthenticated access to protected route redirects to landing', async () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/summary']}>
        <Layout><AppRoutes /></Layout>
      </MemoryRouter>
    );
    expect(await screen.findByText(/Sign in to get started/i)).toBeInTheDocument();
  });

  test('After login, visiting /about redirects to dashboard', async () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/about']}>
        <Layout><AppRoutes /></Layout>
      </MemoryRouter>,
      { account: { name: 'Mock User', username: 'mock@sunbeth.com' } }
    );
    // Dashboard welcome text should appear post-redirect
    await waitFor(async () => {
      expect(await screen.findByText(/Welcome/i)).toBeInTheDocument();
    });
  });

  test('Reader hides Accept for already acknowledged document', async () => {
    renderWithProviders(
      <MemoryRouter initialEntries={['/document/d1?batchId=1']}>
        <Layout>
          <Routes>
            <Route path="/document/:id" element={<AppRoutes />} />
          </Routes>
          <AppRoutes />
        </Layout>
      </MemoryRouter>,
      { account: { name: 'Mock User', username: 'mock@sunbeth.com' } }
    );
    // Accept CTA should not be present for acknowledged doc id
    await waitFor(() => {
      const acceptBtn = screen.queryByRole('button', { name: /I Accept/i });
      expect(acceptBtn).not.toBeInTheDocument();
    });
  });
});
