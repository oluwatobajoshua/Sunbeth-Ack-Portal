import { getApiBase } from './runtimeConfig';
import { showPdfPreview } from './showPdfPreview';
import { tripleDialog, confirmDialog } from './alerts';
import { requestConsentIfNeeded } from './legalConsent';

export type DuePolicy = {
  policyId: number;
  name: string;
  description?: string | null;
  required: boolean;
  fileId: number;
  frequency: string;
  intervalDays?: number | null;
  lastAck?: string | null;
  nextDue: string;
  graceUntil: string;
  overdue: boolean;
};

export async function fetchDuePolicies(email: string): Promise<DuePolicy[]> {
  const base = getApiBase();
  if (!base) return [];
  try {
    const res = await fetch(`${base}/api/policies/due?email=${encodeURIComponent(email)}`, { cache: 'no-store' });
    const j = await res.json().catch(() => null);
    const arr: any[] = Array.isArray(j?.due) ? j.due : [];
    return arr as DuePolicy[];
  } catch {
    return [];
  }
}

async function acknowledgePolicy(email: string, fileId: number): Promise<boolean> {
  const base = getApiBase();
  if (!base) return false;
  try {
    const res = await fetch(`${base}/api/policies/ack`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, fileId })
    });
    if (res.status === 403) {
      // Backend requires legal consent first
      return false;
    }
    return res.ok;
  } catch {
    return false;
  }
}

// Prompt the user to complete due policies. Returns true if all completed or none due.
export async function enforceDuePolicies(email: string): Promise<boolean> {
  const due = await fetchDuePolicies(email);
  if (due.length === 0) return true;
  const base = getApiBase() || '';

  for (const p of due) {
    const url = `${base}/api/files/${p.fileId}`;
    const html = `
      <div style="text-align:left">
        <div style="font-weight:600">${p.name}</div>
        ${p.description ? `<div class="small" style="margin-top:4px">${p.description}</div>` : ''}
        <div class="small muted" style="margin-top:8px">Due: ${new Date(p.nextDue).toLocaleString()} • Grace until: ${new Date(p.graceUntil).toLocaleString()}</div>
        <div class="small muted">Frequency: ${p.frequency}${p.intervalDays ? ` (${p.intervalDays} days)` : ''}</div>
      </div>
    `;
    // offer preview loop
    for (;;) {
      const r = await tripleDialog('Policy acknowledgement required', html, 'Acknowledge', 'Later', 'Preview PDF', { icon: p.overdue ? 'warning' : 'info' });
      if (r === 'deny') break; // Later — stop enforcing for now
      if (r === 'cancel') {
        await showPdfPreview({ title: p.name, url });
        continue;
      }
      // Confirm acknowledge
      const ok = await confirmDialog('Confirm acknowledgement', 'I have read and understand this policy.', 'Confirm', 'Back', { icon: 'question' });
      if (!ok) continue;
      // Ensure legal consent first (policy scope)
      const consented = await requestConsentIfNeeded(email, 'policy');
      if (!consented) {
        // User denied consent; cannot proceed
        break;
      }
      let saved = await acknowledgePolicy(email, p.fileId);
      if (!saved) {
        // Try to recover if backend enforced consent after our check
        const consented2 = await requestConsentIfNeeded(email, 'policy');
        if (consented2) saved = await acknowledgePolicy(email, p.fileId);
      }
      if (!saved) {
        // Give the user a chance to retry or view in new tab
        const retry = await confirmDialog('Save failed', 'We could not record your acknowledgement. Try again?', 'Retry', 'Cancel', { icon: 'error' });
        if (retry) continue;
        break;
      }
      break;
    }
  }

  // Re-check to see if any remain due
  const remain = await fetchDuePolicies(email);
  return remain.length === 0;
}

export default { fetchDuePolicies, enforceDuePolicies };