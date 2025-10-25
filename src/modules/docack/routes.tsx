import React from 'react';
export const docackRoutes = [
  { path: '/admin/docack', element: React.createElement(React.lazy(() => import('./Admin'))) }
];
