export interface MockUser {
  id?: string;
  displayName: string;
  email: string;
  photoUrl?: string;
  roles?: string[];
}

const KEY = 'mock_user_profile';

const defaultUser: MockUser = {
  id: 'mock',
  displayName: 'Mock User',
  email: 'mock.user@sunbeth.com',
  photoUrl: '/logo192.png',
  roles: []
};

export const getCurrentUser = async (): Promise<MockUser> => {
  await new Promise(r => setTimeout(r, 50));
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return defaultUser;
    return JSON.parse(raw) as MockUser;
  } catch (e) {
    try { localStorage.removeItem(KEY); } catch {};
    return defaultUser;
  }
};

export const setMockUser = async (u: Partial<MockUser>): Promise<void> => {
  const cur = (await getCurrentUser()) || defaultUser;
  const next = { ...cur, ...u };
  localStorage.setItem(KEY, JSON.stringify(next));
  try { window.dispatchEvent(new CustomEvent('mockUserChanged', { detail: next })); } catch (e) {}
};

export const clearMockUser = async (): Promise<void> => {
  localStorage.removeItem(KEY);
  try { window.dispatchEvent(new CustomEvent('mockUserChanged', { detail: null })); } catch (e) {}
};

export const seedMockUser = async (u?: MockUser) => {
  localStorage.setItem(KEY, JSON.stringify(u || defaultUser));
  try { window.dispatchEvent(new CustomEvent('mockUserChanged', { detail: u || defaultUser })); } catch (e) {}
};
