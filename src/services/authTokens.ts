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

// Acquire a token for Dataverse resource using /.default scope for the org URL
export const getDataverseToken = async (): Promise<string> => {
  const orgUrl = (process.env.REACT_APP_DATAVERSE_URL || '').replace(/\/$/, '');
  if (!orgUrl) throw new Error('REACT_APP_DATAVERSE_URL not configured');
  // Use /.default to leverage app registration delegated permissions for Dataverse
  const scope = `${orgUrl}/.default`;
  const acct = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!acct) throw new Error('No active account for Dataverse token');
  try {
    const resp = await msalInstance.acquireTokenSilent({ account: acct, scopes: [scope] });
    return resp.accessToken;
  } catch {
    const mode = getInteractiveMode();
    if (mode === 'popup') {
      const resp = await msalInstance.acquireTokenPopup({ scopes: [scope] });
      return resp.accessToken;
    } else {
      await msalInstance.acquireTokenRedirect({ scopes: [scope] });
      throw new Error('Interactive redirect initiated for Dataverse token');
    }
  }
};
