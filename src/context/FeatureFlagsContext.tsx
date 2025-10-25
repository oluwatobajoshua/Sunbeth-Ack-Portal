import React, { createContext, useContext, useEffect, useState } from 'react';
import { apiGet } from '../services/api';

type FlagsContext = {
  externalSupport: boolean;
  flags: Record<string, boolean>;
  loaded: boolean;
  refresh: () => void;
};

const FeatureFlagsCtx = createContext<FlagsContext>({ externalSupport: false, flags: {}, loaded: false, refresh: () => {} });

export const FeatureFlagsProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [externalSupport, setExternalSupport] = useState(false);
  const [flags, setFlags] = useState<Record<string, boolean>>({});
  const [loaded, setLoaded] = useState(false);

  const load = async () => {
    setLoaded(false);
    try {
      const [ext, eff] = await Promise.allSettled([
        apiGet('/api/settings/external-support'),
        apiGet('/api/flags/effective')
      ]);
      if (ext.status === 'fulfilled') {
        setExternalSupport(!!(ext.value as any)?.enabled);
      } else {
        setExternalSupport(false);
      }
      if (eff.status === 'fulfilled') {
        const f = (eff.value as any)?.flags;
        setFlags(f && typeof f === 'object' ? f as Record<string, boolean> : {});
      } else {
        setFlags({});
      }
    } catch {
      setExternalSupport(false);
      setFlags({});
    } finally {
      setLoaded(true);
    }
  };

  useEffect(() => { load(); }, []);

  return (
    <FeatureFlagsCtx.Provider value={{ externalSupport, flags, loaded, refresh: load }}>
      {children}
    </FeatureFlagsCtx.Provider>
  );
};

export const useFeatureFlags = () => useContext(FeatureFlagsCtx);
