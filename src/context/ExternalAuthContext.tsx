import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';

export type ExternalUser = { email: string; name?: string | null };

type ExternalAuth = {
  user: ExternalUser | null;
  isAuthenticated: boolean;
  login: (user: ExternalUser) => void;
  logout: () => void;
};

const defaultValue: ExternalAuth = {
  user: null,
  isAuthenticated: false,
  login: () => {},
  logout: () => {}
};

const ExternalAuthContext = createContext<ExternalAuth>(defaultValue);
export const useExternalAuth = () => useContext(ExternalAuthContext);

const STORAGE_KEY = 'sunbeth:externalAuth';

export const ExternalAuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<ExternalUser | null>(null);

  // Load from localStorage
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed.email === 'string') {
          setUser({ email: parsed.email, name: parsed.name ?? null });
        }
      }
    } catch {}
  }, []);

  // Persist to localStorage
  useEffect(() => {
    try {
      if (user) localStorage.setItem(STORAGE_KEY, JSON.stringify(user));
      else localStorage.removeItem(STORAGE_KEY);
    } catch {}
  }, [user]);

  const login = (u: ExternalUser) => setUser(u);
  const logout = () => setUser(null);

  const value = useMemo<ExternalAuth>(() => ({
    user,
    isAuthenticated: !!user,
    login,
    logout
  }), [user]);

  return (
    <ExternalAuthContext.Provider value={value}>
      {children}
    </ExternalAuthContext.Provider>
  );
};
