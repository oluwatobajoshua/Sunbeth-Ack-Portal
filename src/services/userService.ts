import { useRuntimeMock } from '../utils/runtimeMock';
import { getCurrentUser as getMockUser, setMockUser, clearMockUser, type MockUser } from './mockUserService';
import { msalInstance } from './msalConfig';

export type { MockUser };

const useMock = () => useRuntimeMock();

export const getCurrentUser = async (): Promise<MockUser> => {
  if (useMock()) return getMockUser();
  // Live: derive from MSAL account if present
  const acct = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!acct) {
    // anonymous until login; keep structure compatible
    return { id: undefined, displayName: 'Guest', email: '', photoUrl: undefined, roles: [] };
  }
  return {
    id: acct.homeAccountId,
    displayName: (acct as any).name || acct.username || 'User',
    email: acct.username,
    photoUrl: undefined,
    roles: []
  };
};

export const updateMockUser = async (u: Partial<MockUser>) => {
  if (useMock()) return setMockUser(u);
  // no-op for live
};

export const clearMockUserProfile = async () => {
  if (useMock()) return clearMockUser();
  // no-op for live
};
