/**
 * Central runtime configuration helpers for environment-based mode.
 *
 * Definitions:
 * - Mock mode: REACT_APP_USE_MOCK === 'true' (loads from in-browser mocks)
 * - Dataverse enabled: REACT_APP_ENABLE_DATAVERSE === 'true' and URL is set
 * - Live mode: not mock AND Dataverse enabled
 */
export const isMockMode = (): boolean => (process.env.REACT_APP_USE_MOCK === 'true');

export const isDataverseEnabled = (): boolean => (
  process.env.REACT_APP_ENABLE_DATAVERSE === 'true' && !!process.env.REACT_APP_DATAVERSE_URL
);

export const isLiveMode = (): boolean => (!isMockMode() && isDataverseEnabled());
