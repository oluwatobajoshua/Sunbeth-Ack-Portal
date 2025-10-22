import { msalInstance } from '../services/msalConfig';
import { info, warn, error as logError } from './logger';
import { getGraphToken, getDataverseToken } from '../services/authTokens';

export type Step = { name: string; ok: boolean; detail?: string };

const baseUrl = `${process.env.REACT_APP_DATAVERSE_URL?.replace(/\/$/, '')}/api/data/v9.2`;

export async function runAuthAndGraphCheck(): Promise<Step[]> {
  const steps: Step[] = [];
  try {
    // env checks
    const clientId = process.env.REACT_APP_CLIENT_ID;
    const tenant = process.env.REACT_APP_TENANT_ID;
  const dv = process.env.REACT_APP_DATAVERSE_URL;
  const dvEnabled = process.env.REACT_APP_ENABLE_DATAVERSE === 'true';
    steps.push({ name: 'Env: REACT_APP_CLIENT_ID', ok: !!clientId, detail: clientId ? clientId : 'MISSING' });
    steps.push({ name: 'Env: REACT_APP_TENANT_ID', ok: !!tenant, detail: tenant ? tenant : 'MISSING' });
  steps.push({ name: 'Env: REACT_APP_DATAVERSE_URL', ok: !!dv, detail: dv ? dv : 'MISSING (skip if Dataverse disabled)' });
  steps.push({ name: 'Env: REACT_APP_ENABLE_DATAVERSE', ok: true, detail: dvEnabled ? 'true (enabled)' : 'false (disabled - Dataverse will be skipped)' });

    // accounts
    const accts = msalInstance.getAllAccounts();
    steps.push({ name: 'MSAL: getAllAccounts', ok: Array.isArray(accts) && accts.length > 0, detail: `accounts=${accts.length}` });

    let account = accts && accts[0];
    if (!account) {
      steps.push({ name: 'MSAL: No signed-in account', ok: false, detail: 'No account found. Please sign in or use mock mode.' });
      return steps;
    }

    steps.push({ name: 'MSAL: account', ok: true, detail: account.username || account.homeAccountId || 'unknown' });

    // Acquire Graph token via helper
    let graphToken: string | null = null;
    try {
      graphToken = await getGraphToken(['User.Read']);
      steps.push({ name: 'Graph token', ok: true, detail: `token length=${graphToken.length}` });
    } catch (e) {
      steps.push({ name: 'Graph token', ok: false, detail: String(e) });
    }

    // Graph: /me
    try {
      const res = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { Authorization: `Bearer ${graphToken}` } });
      if (!res.ok) {
        steps.push({ name: 'Graph: GET /me', ok: false, detail: `status=${res.status}` });
      } else {
        const j = await res.json();
        steps.push({ name: 'Graph: GET /me', ok: true, detail: `id=${j.id} displayName=${j.displayName}` });
      }
    } catch (e) {
      logError('Graph /me fetch failed', e);
      steps.push({ name: 'Graph: GET /me', ok: false, detail: String(e) });
    }

    // Graph: groups listing (Group.Read.All)
    try {
      const grpTok = await getGraphToken(['Group.Read.All']);
      const res = await fetch('https://graph.microsoft.com/v1.0/groups?$top=1', { headers: { Authorization: `Bearer ${grpTok}` } });
      steps.push({ name: 'Graph: GET /groups', ok: res.ok, detail: res.ok ? 'ok' : `status=${res.status}` });
    } catch (e) {
      steps.push({ name: 'Graph: GET /groups', ok: false, detail: 'Group.Read.All not consented?' });
    }

    // Graph: users listing (User.Read.All)
    try {
      const usrTok = await getGraphToken(['User.Read.All']);
      const res = await fetch('https://graph.microsoft.com/v1.0/users?$top=1', { headers: { Authorization: `Bearer ${usrTok}` } });
      steps.push({ name: 'Graph: GET /users', ok: res.ok, detail: res.ok ? 'ok' : `status=${res.status}` });
    } catch (e) {
      steps.push({ name: 'Graph: GET /users', ok: false, detail: 'User.Read.All not consented?' });
    }

    // SharePoint: sites and drives (Sites.Read.All + Files.Read.All)
    try {
      const spTok = await getGraphToken(['Sites.Read.All','Files.Read.All']);
      const sres = await fetch('https://graph.microsoft.com/v1.0/sites?search=*', { headers: { Authorization: `Bearer ${spTok}` } });
      const okSites = sres.ok;
      steps.push({ name: 'Graph: GET /sites', ok: okSites, detail: okSites ? 'ok' : `status=${sres.status}` });
      if (okSites) {
        try {
          const js = await sres.json();
          const first = js?.value?.[0]?.id;
          if (first) {
            const dres = await fetch(`https://graph.microsoft.com/v1.0/sites/${first}/drives?$top=1`, { headers: { Authorization: `Bearer ${spTok}` } });
            steps.push({ name: 'Graph: GET site drives', ok: dres.ok, detail: dres.ok ? 'ok' : `status=${dres.status}` });
          }
        } catch {}
      }
    } catch (e) {
      steps.push({ name: 'Graph: SharePoint access', ok: false, detail: 'Sites.Read.All and Files.Read.All not consented?' });
    }
    // SharePoint upload requires Files.ReadWrite.All explicitly; check token acquisition
    try {
      const upTok = await getGraphToken(['Files.ReadWrite.All','Sites.Read.All']);
      steps.push({ name: 'Graph: Upload scope check', ok: !!upTok, detail: 'Files.ReadWrite.All acquired' });
    } catch {
      steps.push({ name: 'Graph: Upload scope check', ok: false, detail: 'Files.ReadWrite.All not consented?' });
    }

    // Mail send check
    try {
      const mailTok = await getGraphToken(['Mail.Send']);
      steps.push({ name: 'Graph: Mail.Send scope', ok: !!mailTok, detail: 'Mail.Send acquired' });
    } catch {
      steps.push({ name: 'Graph: Mail.Send scope', ok: false, detail: 'Mail.Send not consented?' });
    }

    // Dataverse: basic call to toba_batches (only if explicitly enabled)
    if (dvEnabled && dv) {
      try {
        // Use Dataverse token (/.default)
        const dvToken = await getDataverseToken();
        const res = await fetch(`${baseUrl}/toba_batches?$top=1`, { headers: { Authorization: `Bearer ${dvToken}`, Accept: 'application/json' } });
        if (!res.ok) {
          const detail = res.status === 404
            ? '404 Not Found: Ensure a Dataverse database exists for this environment.'
            : `status=${res.status}`;
          steps.push({ name: 'Dataverse: GET toba_batches', ok: false, detail });
        } else {
          const j = await res.json();
          const c = Array.isArray(j.value) ? j.value.length : 0;
          steps.push({ name: 'Dataverse: GET toba_batches', ok: true, detail: `count=${c}` });
        }
      } catch (e) {
        logError('Dataverse fetch failed', e);
        const msg = String(e || '')
        const hint = msg.includes('AADSTS650057')
          ? 'Invalid resource: Add Dynamics CRM (Dataverse) delegated permission "user_impersonation" to your app registration and grant admin consent, or disable Dataverse by setting REACT_APP_ENABLE_DATAVERSE=false.'
          : 'If you are not using Dataverse, set REACT_APP_ENABLE_DATAVERSE=false. If you intend to use it, create a Dataverse database for this environment (see dataverse/scripts/create-database.ps1).';
        steps.push({ name: 'Dataverse: GET toba_batches', ok: false, detail: `${msg}\n${hint}` });
      }
    } else {
      steps.push({ name: 'Dataverse: Skipped', ok: true, detail: 'Disabled or not configured' });
    }

  } catch (err) {
    logError('Health check failed', err);
    steps.push({ name: 'Health check', ok: false, detail: String(err) });
  }
  try {
    // Append quick guidance at the end
    const missing = steps.filter(s => !s.ok && s.name.startsWith('Graph'));
    if (missing.length) {
      const needed = missing.map(s => s.name.includes('groups') ? 'Group.Read.All'
        : s.name.includes('/users') ? 'User.Read.All'
        : s.name.includes('Upload') ? 'Files.ReadWrite.All'
        : s.name.includes('/sites') || s.name.includes('drives') ? 'Sites.Read.All'
        : s.name.includes('Mail.Send') ? 'Mail.Send'
        : 'User.Read').filter((v, i, a) => a.indexOf(v) === i);
      steps.push({ name: 'Grant these scopes', ok: false, detail: needed.join(', ') });
    }
  } catch {}
  return steps;
}

// expose on window for manual invocation in dev
try { (window as any).__sunbethRunDiagnostics = runAuthAndGraphCheck } catch { }

