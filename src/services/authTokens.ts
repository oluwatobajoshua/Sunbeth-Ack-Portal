import { msalInstance } from './msalConfig';
import { getInteractiveMode } from './authInteractive';

// Acquire a token for Microsoft Graph with the given scopes (delegated)
export const getGraphToken = async (scopes: string[] = ['User.Read']): Promise<string> => {
  const acct = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!acct) throw new Error('No active account for token acquisition');
  try {
    const resp = await msalInstance.acquireTokenSilent({ account: acct, scopes });
    return resp.accessToken;
  } catch {
    const mode = getInteractiveMode();
    if (mode === 'popup') {
      const resp = await msalInstance.acquireTokenPopup({ scopes });
      return resp.accessToken;
    } else {
      await msalInstance.acquireTokenRedirect({ scopes });
      throw new Error('Interactive redirect initiated for Graph token');
    }
  }
};


