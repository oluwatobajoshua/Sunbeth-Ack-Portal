import React, { createContext, useContext, useEffect, useState } from 'react';
import { getApiBase } from '../utils/runtimeConfig';

type Flags = {
  externalSupport: boolean;
  loaded: boolean;
  refresh: () => void;
};

const FeatureFlagsCtx = createContext<Flags>({ externalSupport: false, loaded: false, refresh: () => {} });

export const FeatureFlagsProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [externalSupport, setExternalSupport] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const apiBase = (getApiBase() as string) || '';

  const load = async () => {
    setLoaded(false);
    try {
      if (!apiBase) { setExternalSupport(false); setLoaded(true); return; }
      const res = await fetch(`${apiBase}/api/settings/external-support`, { cache: 'no-store' });
      if (!res.ok) throw new Error('settings_failed');
      const j = await res.json();
      setExternalSupport(!!j?.enabled);
    } catch {
      setExternalSupport(false);
    } finally {
      setLoaded(true);
    }
  };

  useEffect(() => { load(); }, [apiBase]);

  return (
    <FeatureFlagsCtx.Provider value={{ externalSupport, loaded, refresh: load }}>
      {children}
    </FeatureFlagsCtx.Provider>
  );
};

export const useFeatureFlags = () => useContext(FeatureFlagsCtx);
