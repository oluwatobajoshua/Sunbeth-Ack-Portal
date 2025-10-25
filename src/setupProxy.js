/*
  CRA dev server proxy for local Express API.
  - Proxies same-origin /api/* requests to the local backend during `npm start`

  Environment knobs:
  - REACT_APP_DEV_API_TARGET: Full target base. If set, we use it verbatim.
    Example: http://127.0.0.1:4000
  - REACT_APP_DEV_TENANT_DOMAIN: Optional tenant domain header to simulate multi-tenant in dev
    Example: orga.local.test (sent as X-Tenant-Domain to backend)
*/

const { createProxyMiddleware } = require('http-proxy-middleware');

module.exports = function (app) {
  const target = process.env.REACT_APP_DEV_API_TARGET || 'http://127.0.0.1:4000';
  const tenantDomain = process.env.REACT_APP_DEV_TENANT_DOMAIN || '';

  // Example mapping:
  //   /api/proxy -> http://127.0.0.1:4000/api/proxy
  app.use(
    '/api',
    createProxyMiddleware({
      target,
      changeOrigin: true,
      ws: false,
      logLevel: 'silent',
      // Keep the /api prefix so paths are forwarded unchanged
      onProxyReq: (proxyReq) => {
        // Ensure no caching during dev
        proxyReq.setHeader('Cache-Control', 'no-cache');
        // Inject tenant header if provided to emulate multi-tenant routing without hosts changes
        if (tenantDomain) {
          try { proxyReq.setHeader('X-Tenant-Domain', tenantDomain); } catch {}
        }
      }
    })
  );
};
