/**
 * Returns whether the application should use mock data.
 * Controlled exclusively by the build-time environment variable REACT_APP_USE_MOCK.
 * Note: Changing this requires restarting the dev server.
 */
export const useRuntimeMock = () => {
  return process.env.REACT_APP_USE_MOCK === 'true';
};
