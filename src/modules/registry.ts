// Registry for dynamically added admin modules.
// The scaffolder will add imports here over time.

export type AdminModule = {
  name: string;
  title: string;
  adminRoute: string; // e.g., /admin/invoicing
  featureFlag?: string; // e.g., module_invoicing_enabled
};

// Initially empty; modules can be appended by scaffolder or manually.
export const adminModules: AdminModule[] = [];
// Core module: Document Acknowledgement
adminModules.push({ name: 'docack', title: 'Document Acknowledgement', adminRoute: '/admin/docack', featureFlag: 'module_docack_enabled' });
