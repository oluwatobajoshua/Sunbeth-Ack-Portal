// Central configuration for Dataverse entity set names and attribute names.
// Allows aligning the app with a pre-existing Power Platform solution without code edits.

export const DV_SETS = {
  batchesSet: process.env.REACT_APP_DV_BATCHES_SET || 'toba_batches',
  documentsSet: process.env.REACT_APP_DV_DOCUMENTS_SET || 'toba_documents',
  userAcksSet: process.env.REACT_APP_DV_USERACKS_SET || 'toba_useracknowledgements',
  userProgressesSet: process.env.REACT_APP_DV_USERPROGRESSES_SET || 'toba_batchuserprogress',
  // New: Businesses and Batch Recipients (for per-user business assignment per batch)
  businessesSet: process.env.REACT_APP_DV_BUSINESSES_SET || 'toba_businesses',
  batchRecipientsSet: process.env.REACT_APP_DV_BATCHRECIPIENTS_SET || 'toba_batchrecipients'
};

export const DV_ATTRS = {
  // Lookup attribute logical names used when creating related records
  documentBatchLookup: process.env.REACT_APP_DV_DOC_BATCH_LOOKUP || 'toba_Batch',
  // Document attribute names (optional; fallback used if undefined)
  docTitleField: process.env.REACT_APP_DV_DOC_TITLE_FIELD || '',
  docUrlField: process.env.REACT_APP_DV_DOC_URL_FIELD || '',
  docVersionField: process.env.REACT_APP_DV_DOC_VERSION_FIELD || '',
  docRequiresSigField: process.env.REACT_APP_DV_DOC_REQUIRES_SIG_FIELD || '',
  ackBatchLookup: process.env.REACT_APP_DV_ACK_BATCH_LOOKUP || 'toba_Batch',
  ackDocumentLookup: process.env.REACT_APP_DV_ACK_DOC_LOOKUP || 'toba_Document',
  ackUserField: process.env.REACT_APP_DV_ACK_USER_FIELD || 'toba_User',
  // New: BatchRecipient field and lookups
  batchRecipientBatchLookup: process.env.REACT_APP_DV_BR_BATCH_LOOKUP || 'toba_Batch',
  batchRecipientBusinessLookup: process.env.REACT_APP_DV_BR_BUSINESS_LOOKUP || 'toba_Business',
  batchRecipientUserField: process.env.REACT_APP_DV_BR_USER_FIELD || '',
  batchRecipientEmailField: process.env.REACT_APP_DV_BR_EMAIL_FIELD || 'toba_Email',
  batchRecipientDisplayNameField: process.env.REACT_APP_DV_BR_DISPLAYNAME_FIELD || 'toba_DisplayName',
  batchRecipientDepartmentField: process.env.REACT_APP_DV_BR_DEPARTMENT_FIELD || 'toba_Department',
  batchRecipientJobTitleField: process.env.REACT_APP_DV_BR_JOBTITLE_FIELD || 'toba_JobTitle',
  batchRecipientLocationField: process.env.REACT_APP_DV_BR_LOCATION_FIELD || 'toba_Location',
  batchRecipientPrimaryGroupField: process.env.REACT_APP_DV_BR_PRIMARYGROUP_FIELD || 'toba_PrimaryGroup'
};

// Feature flags for Dataverse behaviors
export const DV_FLAGS = {
  // When true, if reading Batch Recipients returns 401/403, we will fall back to listing all batches
  // so the Dashboard isn't empty during permission rollout. Default: false (secure by default).
  fallbackAllBatchesOn401: (process.env.REACT_APP_DV_FALLBACK_ALL_BATCHES_ON_401 || 'false').toLowerCase() === 'true'
};
