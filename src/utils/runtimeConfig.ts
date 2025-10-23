/**
 * Central runtime configuration helpers for environment-based mode.
 *
 * Definitions:
 * - SQLite mode: REACT_APP_ENABLE_SQLITE === 'true' and API base is set
 * - SharePoint Lists mode: REACT_APP_ENABLE_SP_LISTS === 'true' and site ID is set
 */

export const isSQLiteEnabled = (): boolean => (
  process.env.REACT_APP_ENABLE_SQLITE === 'true' && !!process.env.REACT_APP_API_BASE
);

export const isSharePointListsEnabled = (): boolean => (
  process.env.REACT_APP_ENABLE_SP_LISTS === 'true' && !!process.env.REACT_APP_SP_SITE_ID
);

// Branding helpers
export const getBrandName = (): string => (
  (process.env.REACT_APP_BRAND_NAME || 'Sunbeth')
);

export const getBrandLogoUrl = (): string | undefined => (
  process.env.REACT_APP_BRAND_LOGO_URL || undefined
);

export const getBrandPrimaryColor = (): string => (
  process.env.REACT_APP_BRAND_COLOR || '#5a189a'
);

// HR notification recipients (comma-separated emails)
export const getHrEmails = (): string[] => {
  const raw = (process.env.REACT_APP_HR_EMAILS || '').trim();
  if (!raw) return [];
  return raw.split(',').map(s => s.trim()).filter(Boolean);
};

// Admin notification recipients (comma-separated emails)
export const getAdminEmails = (): string[] => {
  const raw = (process.env.REACT_APP_ADMIN_EMAILS || '').trim();
  if (!raw) return [];
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
};

// Manager notification recipients (comma-separated emails)
export const getManagerEmails = (): string[] => {
  const raw = (process.env.REACT_APP_MANAGER_EMAILS || '').trim();
  if (!raw) return [];
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
};

// Optional CC/BCC lists for completion notifications
export const getCompletionCcEmails = (): string[] => {
  const raw = (process.env.REACT_APP_COMPLETION_CC || '').trim();
  if (!raw) return [];
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
};

export const getCompletionBccEmails = (): string[] => {
  const raw = (process.env.REACT_APP_COMPLETION_BCC || '').trim();
  if (!raw) return [];
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
};

// Core endpoints and flags
export const getApiBase = (): string | null => {
  const base = (process.env.REACT_APP_API_BASE || '').trim();
  if (!base) return null;
  return base.replace(/\/$/, '');
};

export const getFlowAckUrl = (): string | null => {
  const url = (process.env.REACT_APP_FLOW_CREATE_USER_ACK_URL || '').trim();
  return url || null;
};

export const isAdminLight = (): boolean => (
  (process.env.REACT_APP_ADMIN_LIGHT || '').toLowerCase() === 'true'
);

export const useAdminModalSelectors = (): boolean => (
  (process.env.REACT_APP_ADMIN_MODAL_SELECTORS || '').toLowerCase() === 'true'
);

// Auth
export const getClientId = (): string | null => (process.env.REACT_APP_CLIENT_ID || null);
export const getTenantId = (): string | null => (process.env.REACT_APP_TENANT_ID || null);

// Busy overlay timing configuration
// You can set either milliseconds or seconds; seconds take precedence if present.
export const getBusyOverlayShowDelayMs = (): number => {
  const secs = Number(process.env.REACT_APP_BUSY_DELAY_SECS);
  if (Number.isFinite(secs) && secs >= 0) return Math.round(secs * 1000);
  const ms = Number(process.env.REACT_APP_BUSY_DELAY_MS);
  if (Number.isFinite(ms) && ms >= 0) return Math.round(ms);
  return 150; // default
};

export const getBusyOverlayMinVisibleMs = (): number => {
  const secs = Number(process.env.REACT_APP_BUSY_MIN_VISIBLE_SECS);
  if (Number.isFinite(secs) && secs >= 0) return Math.round(secs * 1000);
  const ms = Number(process.env.REACT_APP_BUSY_MIN_VISIBLE_MS);
  if (Number.isFinite(ms) && ms >= 0) return Math.round(ms);
  return 600; // default
};
