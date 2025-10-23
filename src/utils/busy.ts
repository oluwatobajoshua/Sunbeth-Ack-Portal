/**
 * Lightweight global busy indicator controller.
 * Use busyPush()/busyPop() to show/hide the overlay.
 */
export const busyPush = (label?: string) => {
  try { window.dispatchEvent(new CustomEvent('sunbeth:busy:push', { detail: { label: label || '' } })); } catch {}
};

export const busyPop = () => {
  try { window.dispatchEvent(new CustomEvent('sunbeth:busy:pop')); } catch {}
};

export const busyReset = () => {
  try { window.dispatchEvent(new CustomEvent('sunbeth:busy:reset')); } catch {}
};
