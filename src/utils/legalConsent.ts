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

async function sendConsentReceipt(userEmail?: string | null, batchId?: string | null) {
  try {
    const base = getApiBase() || '';
    if (!base || !userEmail) return;
    await fetch(`${base}/api/consents`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email: userEmail, batchId: batchId || undefined })
    });
  } catch (e) {
    // non-blocking; log for diagnostics only
    console.debug('consent receipt failed (non-blocking)', e);
  }
}

// eslint-disable-next-line complexity, max-lines-per-function
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
      <p>By clicking “I Agree”, I hereby acknowledge and confirm that:</p>
      <ul style="margin-left:1em">
        <li>I have read and understood the content of the document presented to me;</li>
        <li>I am signing/acknowledging this document voluntarily and without coercion;</li>
        <li>The information I have provided (if any) is true and accurate to the best of my knowledge;</li>
        <li>My electronic acknowledgement/signature has the same legal effect as a handwritten signature, to the fullest extent permitted by applicable law; and</li>
        <li>I consent to the electronic delivery, execution, and storage of this document by [App/Company Name].</li>
      </ul>
      ${previewUrl ? '<p style="margin:8px 0">You can preview the official PDF inside the app before agreeing.</p>' : ''}
      <p>I understand that this acknowledgement is binding and enforceable.</p>
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
      // Best-effort server receipt for audit
      void sendConsentReceipt(userEmail, batchId);
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
