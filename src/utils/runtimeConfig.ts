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
