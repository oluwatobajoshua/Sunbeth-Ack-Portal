import { confirmDialog, tripleDialog } from './alerts';
import { getTenantId, getApiBase } from './runtimeConfig';
import { showPdfPreview } from './showPdfPreview';

const CONSENT_PREFIX = 'sunbeth:legalConsent:v1';

function keyFor(userEmail?: string | null, batchId?: string | null) {
  const tenant = getTenantId() || 'default';
  const user = (userEmail || 'anon').toLowerCase();
  const batch = batchId || 'none';
  return `${CONSENT_PREFIX}:${tenant}:${user}:${batch}`;
}

export function hasConsent(userEmail?: string | null, batchId?: string | null): boolean {
  try {
    const k = keyFor(userEmail, batchId);
    return localStorage.getItem(k) === 'true';
  } catch {
    return false;
  }
}

export function recordConsent(userEmail?: string | null, batchId?: string | null) {
  try {
    const k = keyFor(userEmail, batchId);
    localStorage.setItem(k, 'true');
  } catch {
    // ignore
  }
}

// eslint-disable-next-line complexity
export async function requestConsentIfNeeded(userEmail?: string | null, batchId?: string | null): Promise<boolean> {
  if (hasConsent(userEmail, batchId)) return true;
  const title = 'Court-certified acknowledgement consent';
  // Try to load current legal document metadata from server
  let previewUrl: string | null = null;
  try {
    const base = getApiBase() || '';
    if (base) {
      const res = await fetch(`${base}/api/settings/legal-consent`, { cache: 'no-store' });
      const j = await res.json().catch(() => null);
      const url = j?.url ? (base + j.url) : null;
      if (url) previewUrl = url;
    }
  } catch { /* non-blocking */ }
  const html = `
    <div style="text-align:left">
      <p>Before you can acknowledge documents in this batch, you must consent to the following:</p>
      <ul style="margin-left:1em">
        <li>Your acknowledgement constitutes a legally-binding record for employment compliance.</li>
        <li>Your name and timestamp will be stored and may be presented in court if required.</li>
        <li>Any misrepresentation is subject to disciplinary action under company policy.</li>
      </ul>
      ${previewUrl ? '<p style="margin:8px 0">You can preview the official PDF inside the app before agreeing.</p>' : ''}
      <p>You may review documents without consenting, but you cannot submit acknowledgements unless you agree.</p>
    </div>
  `;
  // Helper to show dialog with or without preview option
  const ask = async (): Promise<'agree' | 'deny' | 'preview'> => {
    if (previewUrl) {
      const r = await tripleDialog(title, html, 'I Agree', 'Deny', 'Preview PDF', { icon: 'info' });
      if (r === 'confirm') return 'agree';
      if (r === 'deny') return 'preview';
      return 'deny';
    }
    const ok = await confirmDialog(title, html, 'I Agree', 'Deny', { icon: 'info' });
    return ok ? 'agree' : 'deny';
  };

  // If preview chosen, show preview then re-ask until user agrees or denies
  for (;;) {
    const choice = await ask();
    if (choice === 'agree') {
      recordConsent(userEmail, batchId);
      return true;
    }
    if (choice === 'preview' && previewUrl) {
      await showPdfPreview({ title: 'Legal document (PDF)', url: previewUrl });
      continue;
    }
    return false; // deny
  }
}

export default { hasConsent, recordConsent, requestConsentIfNeeded };
