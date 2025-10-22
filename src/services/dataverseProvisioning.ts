import { getDataverseToken } from './authTokens';
import { DV_SETS, DV_ATTRS } from './dataverseConfig';

const lang = 1033;

const labels = (text: string) => ({ LocalizedLabels: [{ Label: text, LanguageCode: lang }] });

const dvBase = () => (process.env.REACT_APP_DATAVERSE_URL || '').replace(/\/$/, '') + '/api/data/v9.2';

async function dvFetch(path: string, token: string, init?: RequestInit) {
  const res = await fetch(`${dvBase()}${path}`, {
    ...init,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json; charset=utf-8',
      'OData-Version': '4.0',
      'OData-MaxVersion': '4.0',
      ...(init?.headers || {})
    }
  });
  return res;
}

// Convert a logical name like "toba_startdate" to a SchemaName like "toba_StartDate"
function toSchemaName(logical: string): string {
  if (!logical) return logical;
  const [prefix, ...restParts] = logical.split('_');
  if (!restParts.length) return logical.charAt(0).toUpperCase() + logical.slice(1);
  const rest = restParts.join('_');
  const pascal = rest
    .split('_')
    .filter(Boolean)
    .map(part => part.charAt(0).toUpperCase() + part.slice(1))
    .join('');
  return `${prefix}_${pascal}`;
}

async function getEntityMetadataId(token: string, logicalName: string): Promise<string | null> {
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${logicalName}')?$select=MetadataId`, token);
  if (res.ok) {
    const j = await res.json();
    return j.MetadataId as string;
  }
  return null;
}

async function ensureEntity(token: string, logicalName: string, displaySingular: string, displayPlural: string, primaryName: string): Promise<string> {
  const existing = await getEntityMetadataId(token, logicalName);
  if (existing) return existing;
  // Use the CreateEntity action so PrimaryAttribute is provided separately and
  // compatible across environments that reject nested PrimaryAttribute on EntityMetadata.
  const payload = {
    Entity: {
      '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata',
      SchemaName: toSchemaName(logicalName),
      DisplayName: labels(displaySingular),
      DisplayCollectionName: labels(displayPlural),
      OwnershipType: 'UserOwned',
      HasActivities: false,
      HasNotes: true
    },
    PrimaryAttribute: {
      '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
      SchemaName: toSchemaName(primaryName),
      RequiredLevel: { Value: 'None' },
      MaxLength: 200,
      FormatName: { Value: 'Text' },
      DisplayName: labels(primaryName)
    }
  } as any;
  const solutionUniqueName = (process.env.REACT_APP_DV_SOLUTION_UNIQUENAME || '').trim();
  const actionPayload = solutionUniqueName ? { ...payload, SolutionUniqueName: solutionUniqueName } : payload;

  // Try multiple compatible creation routes across environments
  const attempts: Array<{ path: string; body: any; note: string }> = [
    { path: '/Microsoft.Dynamics.CRM.CreateEntity', body: actionPayload, note: 'Unbound action (FQ name)' },
    { path: '/CreateEntity', body: actionPayload, note: 'Unbound action (short name)' },
    { path: '/EntityDefinitions', body: { '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata', SchemaName: toSchemaName(logicalName), DisplayName: labels(displaySingular), DisplayCollectionName: labels(displayPlural), OwnershipType: 'UserOwned', HasActivities: false, HasNotes: true, PrimaryNameAttribute: primaryName, PrimaryAttribute: payload.PrimaryAttribute }, note: 'EntityDefinitions with nested PrimaryAttribute' },
    { path: '/EntityDefinitions', body: { '@odata.type': 'Microsoft.Dynamics.CRM.EntityMetadata', SchemaName: toSchemaName(logicalName), DisplayName: labels(displaySingular), DisplayCollectionName: labels(displayPlural), OwnershipType: 'UserOwned', HasActivities: false, HasNotes: true, PrimaryNameAttribute: primaryName }, note: 'EntityDefinitions PrimaryNameAttribute only' }
  ];

  let created = false; let lastErr = '';
  for (const attempt of attempts) {
    const res = await dvFetch(attempt.path, token, { method: 'POST', body: JSON.stringify(attempt.body) });
    if (res.ok) { created = true; break; }
    const t = await res.text().catch(() => '');
    lastErr = `${attempt.note}: ${res.status} ${t}`;
    // Continue to next strategy on 404/400; break for 401/403 to avoid repeated consent prompts
    if (res.status === 401 || res.status === 403) break;
  }
  if (!created) {
    throw new Error(`Failed to create entity ${logicalName}. Attempts exhausted. Last error: ${lastErr}`);
  }
  const id = await getEntityMetadataId(token, logicalName);
  if (!id) throw new Error(`Entity ${logicalName} created but MetadataId not found`);
  return id;
}

async function attributeExists(token: string, entityLogical: string, attrLogical: string): Promise<boolean> {
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes(LogicalName='${attrLogical}')?$select=MetadataId`, token);
  return res.ok;
}

async function ensureStringAttribute(token: string, entityLogical: string, schemaName: string, displayName: string, maxLength = 4000) {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    MaxLength: maxLength,
    FormatName: { Value: 'Text' }
  };
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create string attribute ${entityLogical}.${schemaName}`);
}

async function ensureIntegerAttribute(token: string, entityLogical: string, schemaName: string, displayName: string) {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.IntegerAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    Format: 'None',
    MinValue: 0,
    MaxValue: 2147483647
  };
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create integer attribute ${entityLogical}.${schemaName}`);
}

async function ensureBooleanAttribute(token: string, entityLogical: string, schemaName: string, displayName: string) {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.BooleanAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    DefaultValue: false,
    OptionSet: {
      '@odata.type': 'Microsoft.Dynamics.CRM.BooleanOptionSetMetadata',
      TrueOption: { Label: labels('Yes'), Value: 1 },
      FalseOption: { Label: labels('No'), Value: 0 }
    }
  } as any;
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create boolean attribute ${entityLogical}.${schemaName}`);
}

async function ensureUrlAttribute(token: string, entityLogical: string, schemaName: string, displayName: string, maxLength = 1000) {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.StringAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    MaxLength: maxLength,
    FormatName: { Value: 'Url' }
  } as any;
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create URL attribute ${entityLogical}.${schemaName}`);
}

async function ensureDateOnlyAttribute(token: string, entityLogical: string, schemaName: string, displayName: string) {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    Format: 'DateOnly',
    DateTimeBehavior: { Value: 'DateOnly' }
  } as any;
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create DateOnly attribute ${entityLogical}.${schemaName}`);
}

async function ensureDateTimeAttribute(token: string, entityLogical: string, schemaName: string, displayName: string, behavior: 'UserLocal' | 'TimeZoneIndependent' = 'UserLocal') {
  if (await attributeExists(token, entityLogical, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    Format: 'DateAndTime',
    DateTimeBehavior: { Value: behavior }
  } as any;
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${entityLogical}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create DateTime attribute ${entityLogical}.${schemaName}`);
}

async function ensureLookupAttribute(token: string, referencingEntity: string, schemaName: string, displayName: string, targetEntity: string) {
  if (await attributeExists(token, referencingEntity, schemaName)) return;
  const body = {
    '@odata.type': 'Microsoft.Dynamics.CRM.LookupAttributeMetadata',
    SchemaName: toSchemaName(schemaName),
    DisplayName: labels(displayName),
    RequiredLevel: { Value: 'None' },
    Targets: [targetEntity]
  } as any;
  const res = await dvFetch(`/EntityDefinitions(LogicalName='${referencingEntity}')/Attributes`, token, { method: 'POST', body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`Failed to create lookup attribute ${referencingEntity}.${schemaName}`);
}

export type ProvisionLog = { step: string; ok: boolean; detail?: string };

// Resolve Dataverse entity set names dynamically by inspecting EntityDefinitions.
// Falls back to DV_SETS env/defaults when detection fails. Useful when publisher prefixes differ.
export async function resolveEntitySets(): Promise<typeof DV_SETS> {
  const token = await getDataverseToken();
  async function detectEntitySetNameBySuffix(suffix: string): Promise<string | undefined> {
    try {
      const res = await dvFetch(`/EntityDefinitions?$select=LogicalName,EntitySetName&$filter=endswith(LogicalName,'%5F${suffix}')`, token, { method: 'GET' });
      if (!res.ok) return undefined;
      const j = await res.json().catch(() => ({ value: [] }));
      const rows: Array<{ LogicalName: string; EntitySetName: string }> = Array.isArray(j?.value) ? j.value : [];
      const preferred = rows.find(r => typeof r?.LogicalName === 'string' && r.LogicalName.toLowerCase().endsWith(`_${suffix.toLowerCase()}`));
      return preferred?.EntitySetName || rows[0]?.EntitySetName;
    } catch { return undefined; }
  }
  const detected = {
    batchesSet: await detectEntitySetNameBySuffix('batch'),
    documentsSet: await detectEntitySetNameBySuffix('batchdocument') || await detectEntitySetNameBySuffix('document'),
    userAcksSet: await detectEntitySetNameBySuffix('batchacknowledgement') || await detectEntitySetNameBySuffix('useracknowledgement'),
    userProgressesSet: await detectEntitySetNameBySuffix('batchuserprogress') || await detectEntitySetNameBySuffix('userprogress'),
    businessesSet: await detectEntitySetNameBySuffix('business'),
    batchRecipientsSet: await detectEntitySetNameBySuffix('batchrecipient')
  } as Partial<typeof DV_SETS>;
  return {
    batchesSet: detected.batchesSet || DV_SETS.batchesSet,
    documentsSet: detected.documentsSet || DV_SETS.documentsSet,
    userAcksSet: detected.userAcksSet || DV_SETS.userAcksSet,
    userProgressesSet: detected.userProgressesSet || DV_SETS.userProgressesSet,
    businessesSet: detected.businessesSet || DV_SETS.businessesSet,
    batchRecipientsSet: detected.batchRecipientsSet || DV_SETS.batchRecipientsSet
  };
}

export async function provisionSunbethSchema(): Promise<ProvisionLog[]> {
  const logs: ProvisionLog[] = [];
  const org = (process.env.REACT_APP_DATAVERSE_URL || '').replace(/\/$/, '');
  if (!org) {
    logs.push({ step: 'Dataverse URL', ok: false, detail: 'REACT_APP_DATAVERSE_URL not set' });
    return logs;
  }
  try {
    const token = await getDataverseToken();
    // Entities (per-entity try/catch with flags)
    let haveBatch = false, haveDoc = false, haveAck = false, haveProg = false, haveBiz = false, haveBR = false;
    try { const id = await ensureEntity(token, 'toba_batch', 'Batch', 'Batches', 'toba_name'); logs.push({ step: 'Entity toba_batch', ok: true, detail: id }); haveBatch = true; } catch (e: any) { logs.push({ step: 'Entity toba_batch', ok: false, detail: String(e?.message || e) }); }
    try { const id = await ensureEntity(token, 'toba_document', 'Document', 'Documents', 'toba_title'); logs.push({ step: 'Entity toba_document', ok: true, detail: id }); haveDoc = true; } catch (e: any) { logs.push({ step: 'Entity toba_document', ok: false, detail: String(e?.message || e) }); }
    try { const id = await ensureEntity(token, 'toba_useracknowledgement', 'User Acknowledgement', 'User Acknowledgements', 'toba_name'); logs.push({ step: 'Entity toba_useracknowledgement', ok: true, detail: id }); haveAck = true; } catch (e: any) { logs.push({ step: 'Entity toba_useracknowledgement', ok: false, detail: String(e?.message || e) }); }
  try { const id = await ensureEntity(token, 'toba_batchuserprogress', 'User Progress', 'User Progresses', 'toba_name'); logs.push({ step: 'Entity toba_batchuserprogress', ok: true, detail: id }); haveProg = true; } catch (e: any) { logs.push({ step: 'Entity toba_batchuserprogress', ok: false, detail: String(e?.message || e) }); }
    try { const id = await ensureEntity(token, 'toba_business', 'Business', 'Businesses', 'toba_name'); logs.push({ step: 'Entity toba_business', ok: true, detail: id }); haveBiz = true; } catch (e: any) { logs.push({ step: 'Entity toba_business', ok: false, detail: String(e?.message || e) }); }
    try { const id = await ensureEntity(token, 'toba_batchrecipient', 'Batch Recipient', 'Batch Recipients', 'toba_name'); logs.push({ step: 'Entity toba_batchrecipient', ok: true, detail: id }); haveBR = true; } catch (e: any) { logs.push({ step: 'Entity toba_batchrecipient', ok: false, detail: String(e?.message || e) }); }

    // Attributes - Batch
    if (haveBatch) {
      await ensureStringAttribute(token, 'toba_batch', 'toba_description', 'Description', 4000);
      await ensureIntegerAttribute(token, 'toba_batch', 'toba_status', 'Status');
      await ensureDateOnlyAttribute(token, 'toba_batch', 'toba_startdate', 'Start Date');
      await ensureDateOnlyAttribute(token, 'toba_batch', 'toba_duedate', 'Due Date');
      logs.push({ step: 'Attributes toba_batch', ok: true });
    } else {
      logs.push({ step: 'Attributes toba_batch', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

    // Attributes - Document
    if (haveDoc) {
      await ensureIntegerAttribute(token, 'toba_document', 'toba_version', 'Version');
      await ensureUrlAttribute(token, 'toba_document', 'toba_fileurl', 'File URL', 1000);
      await ensureBooleanAttribute(token, 'toba_document', 'toba_requiressignature', 'Requires Signature');
      if (haveBatch) await ensureLookupAttribute(token, 'toba_document', 'toba_Batch', 'Batch', 'toba_batch');
      logs.push({ step: 'Attributes toba_document', ok: true });
    } else {
      logs.push({ step: 'Attributes toba_document', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

    // Attributes - User Acknowledgement
    if (haveAck) {
      await ensureBooleanAttribute(token, 'toba_useracknowledgement', 'toba_acknowledged', 'Acknowledged');
      await ensureDateTimeAttribute(token, 'toba_useracknowledgement', 'toba_ackdate', 'Acknowledged On', 'UserLocal');
      if (haveBatch) await ensureLookupAttribute(token, 'toba_useracknowledgement', 'toba_Batch', 'Batch', 'toba_batch');
      if (haveDoc) await ensureLookupAttribute(token, 'toba_useracknowledgement', 'toba_Document', 'Document', 'toba_document');
      await ensureStringAttribute(token, 'toba_useracknowledgement', 'toba_User', 'User', 255);
      logs.push({ step: 'Attributes toba_useracknowledgement', ok: true });
    } else {
      logs.push({ step: 'Attributes toba_useracknowledgement', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

    // Attributes - User Progress
    if (haveProg) {
      await ensureIntegerAttribute(token, 'toba_batchuserprogress', 'toba_acknowledged', 'Acknowledged Count');
      await ensureIntegerAttribute(token, 'toba_batchuserprogress', 'toba_totaldocs', 'Total Documents');
      if (haveBatch) await ensureLookupAttribute(token, 'toba_batchuserprogress', 'toba_Batch', 'Batch', 'toba_batch');
      await ensureStringAttribute(token, 'toba_batchuserprogress', 'toba_User', 'User', 255);
      logs.push({ step: 'Attributes toba_batchuserprogress', ok: true });
    } else {
      logs.push({ step: 'Attributes toba_batchuserprogress', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

    // Attributes - Business
    if (haveBiz) {
      await ensureStringAttribute(token, 'toba_business', 'toba_code', 'Code', 50);
      await ensureStringAttribute(token, 'toba_business', 'toba_description', 'Description', 4000);
      await ensureBooleanAttribute(token, 'toba_business', 'toba_isactive', 'Is Active');
      logs.push({ step: 'Attributes toba_business', ok: true });

      // Seed a few sample businesses if none exist
      try {
        const checkRes = await dvFetch(`/${DV_SETS.businessesSet}?$select=toba_businessid&$top=1`, token, { method: 'GET' });
        if (checkRes.ok) {
          const j = await checkRes.json().catch(() => ({ value: [] }));
          const count = Array.isArray(j?.value) ? j.value.length : 0;
          if (count === 0) {
            const samples = [
              { toba_name: 'Head Office', toba_code: 'HO', toba_isactive: true },
              { toba_name: 'Subsidiary A', toba_code: 'SUB-A', toba_isactive: true },
              { toba_name: 'Subsidiary B', toba_code: 'SUB-B', toba_isactive: true }
            ];
            let seeded = 0;
            for (const s of samples) {
              const r = await dvFetch(`/${DV_SETS.businessesSet}`, token, { method: 'POST', body: JSON.stringify(s) });
              if (r.ok) seeded++;
            }
            logs.push({ step: 'Seed businesses', ok: true, detail: `Inserted ${seeded} sample rows` });
          } else {
            logs.push({ step: 'Seed businesses', ok: true, detail: 'Skipped (records already exist)' });
          }
        } else if (checkRes.status === 404) {
          logs.push({ step: 'Seed businesses', ok: false, detail: 'Skipped (entity set not found)' });
        } else {
          logs.push({ step: 'Seed businesses', ok: false, detail: `Skipped (${checkRes.status})` });
        }
      } catch (e: any) {
        logs.push({ step: 'Seed businesses', ok: false, detail: String(e?.message || e) });
      }
    } else {
      logs.push({ step: 'Attributes toba_business', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

    // Attributes - Batch Recipient
    if (haveBR) {
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_User', 'User', 255);
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_Email', 'Email', 255);
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_DisplayName', 'Display Name', 255);
      if (haveBatch) await ensureLookupAttribute(token, 'toba_batchrecipient', 'toba_Batch', 'Batch', 'toba_batch');
      if (haveBiz) await ensureLookupAttribute(token, 'toba_batchrecipient', 'toba_Business', 'Business', 'toba_business');
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_Department', 'Department', 255);
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_JobTitle', 'Job Title', 255);
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_Location', 'Location', 255);
      await ensureStringAttribute(token, 'toba_batchrecipient', 'toba_PrimaryGroup', 'Primary Group', 255);
      logs.push({ step: 'Attributes toba_batchrecipient', ok: true });
    } else {
      logs.push({ step: 'Attributes toba_batchrecipient', ok: false, detail: 'Skipped (entity not created). Use CSV templates to create the table, then re-run.' });
    }

  } catch (e: any) {
    logs.push({ step: 'Provisioning failed', ok: false, detail: String(e?.message || e) });
  }
  return logs;
}

// Lightweight connectivity check: calls WhoAmI to verify token and org URL
export async function whoAmI(): Promise<{ UserId?: string; OrganizationId?: string }> {
  const token = await getDataverseToken();
  const res = await dvFetch('/WhoAmI', token, { method: 'GET' });
  if (!res.ok) throw new Error(`WhoAmI failed: ${res.status} ${await res.text().catch(() => '')}`);
  return res.json();
}

// Seed minimal sample data across core tables so the Admin can verify DV reads/writes quickly
export async function seedSunbethSampleData(): Promise<ProvisionLog[]> {
  const logs: ProvisionLog[] = [];
  try {
    const token = await getDataverseToken();
    const base = dvBase();
    const SETS = await resolveEntitySets();
    try { logs.push({ step: 'Detect entity sets', ok: true, detail: JSON.stringify(SETS) }); } catch { logs.push({ step: 'Detect entity sets', ok: true }); }

    // 1) Ensure there's at least one Business; create one if needed
    let businessId: string | undefined = undefined;
    try {
      const r = await dvFetch(`/${SETS.businessesSet}?$select=toba_businessid&$top=1`, token, { method: 'GET' });
      if (r.ok) {
        const j = await r.json().catch(() => ({ value: [] }));
        const row = Array.isArray(j?.value) ? j.value[0] : undefined;
        if (row?.toba_businessid) businessId = row.toba_businessid as string;
      }
      if (!businessId) {
        const r2 = await dvFetch(`/${SETS.businessesSet}`, token, { method: 'POST', body: JSON.stringify({ toba_name: 'Head Office', toba_code: 'HO', toba_isactive: true }) });
        if (!r2.ok) throw new Error(`create business: ${r2.status} ${await r2.text().catch(() => '')}`);
        const loc = r2.headers.get('OData-EntityId') || '';
        const m = loc.match(/[0-9a-fA-F-]{36}/);
        businessId = m ? m[0] : undefined;
        logs.push({ step: 'Seed business', ok: true, detail: businessId });
      } else {
        logs.push({ step: 'Seed business', ok: true, detail: 'Skipped (exists)' });
      }
    } catch (e: any) {
      logs.push({ step: 'Seed business', ok: false, detail: String(e?.message || e) });
    }

    // 2) Create a Batch
    let batchId: string | undefined = undefined;
    try {
      const body: any = {
        toba_name: `HR Policies - ${new Date().getFullYear()}`,
        toba_startdate: new Date().toISOString().slice(0,10),
        toba_duedate: new Date(Date.now() + 7*24*60*60*1000).toISOString().slice(0,10),
        toba_description: 'Sample batch for verification',
        toba_status: 1
      };
  const r = await dvFetch(`/${SETS.batchesSet}`, token, { method: 'POST', body: JSON.stringify(body) });
      if (!r.ok) throw new Error(`create batch: ${r.status} ${await r.text().catch(() => '')}`);
      const loc = r.headers.get('OData-EntityId') || '';
      const m = loc.match(/[0-9a-fA-F-]{36}/);
      batchId = m ? m[0] : undefined;
      logs.push({ step: 'Seed batch', ok: true, detail: batchId });
    } catch (e: any) {
      logs.push({ step: 'Seed batch', ok: false, detail: String(e?.message || e) });
    }

    // 3) Create a Document linked to the Batch
    let documentId: string | undefined = undefined;
    try {
      if (!batchId) throw new Error('no batch id');
      const docBody: any = {
        toba_title: 'Code of Conduct.pdf',
        toba_fileurl: 'https://contoso.sharepoint.com/sites/hr/Shared%20Documents/Code%20of%20Conduct.pdf',
        toba_version: 1,
        [`${DV_ATTRS.documentBatchLookup}@odata.bind`]: `/${SETS.batchesSet}(${batchId})`
      };
      const r = await dvFetch(`/${SETS.documentsSet}`, token, { method: 'POST', body: JSON.stringify(docBody) });
      if (!r.ok) throw new Error(`create document: ${r.status} ${await r.text().catch(() => '')}`);
      const loc = r.headers.get('OData-EntityId') || '';
      const m = loc.match(/[0-9a-fA-F-]{36}/);
      documentId = m ? m[0] : undefined;
      logs.push({ step: 'Seed document', ok: true, detail: documentId });
    } catch (e: any) {
      logs.push({ step: 'Seed document', ok: false, detail: String(e?.message || e) });
    }

    // 4) Create a Batch Recipient with Business
    try {
      if (!batchId) throw new Error('no batch id');
      const brBody: any = {
        toba_name: 'Recipient - jane.doe@contoso.com',
        [DV_ATTRS.batchRecipientUserField]: 'jane.doe@contoso.com',
        [DV_ATTRS.batchRecipientEmailField]: 'jane.doe@contoso.com',
        [DV_ATTRS.batchRecipientDisplayNameField]: 'Jane Doe',
        [DV_ATTRS.batchRecipientDepartmentField]: 'HR',
        [DV_ATTRS.batchRecipientJobTitleField]: 'HR Manager',
        [DV_ATTRS.batchRecipientLocationField]: 'Lagos',
        [`${DV_ATTRS.batchRecipientBatchLookup}@odata.bind`]: `/${SETS.batchesSet}(${batchId})`
      };
      if (businessId) brBody[`${DV_ATTRS.batchRecipientBusinessLookup}@odata.bind`] = `/${SETS.businessesSet}(${businessId})`;
      const r = await dvFetch(`/${SETS.batchRecipientsSet}`, token, { method: 'POST', body: JSON.stringify(brBody) });
      if (!r.ok) throw new Error(`create batch recipient: ${r.status} ${await r.text().catch(() => '')}`);
      logs.push({ step: 'Seed batch recipient', ok: true });
    } catch (e: any) {
      logs.push({ step: 'Seed batch recipient', ok: false, detail: String(e?.message || e) });
    }

    // 5) Create a User Progress row (optional)
    try {
      if (!batchId) throw new Error('no batch id');
      const prBody: any = {
        toba_name: 'Progress - jane.doe@contoso.com',
        toba_acknowledged: 0,
        toba_totaldocs: 1,
        [DV_ATTRS.ackUserField]: 'jane.doe@contoso.com',
        [`${DV_ATTRS.ackBatchLookup}@odata.bind`]: `/${SETS.batchesSet}(${batchId})`
      };
      const r = await dvFetch(`/${SETS.userProgressesSet}`, token, { method: 'POST', body: JSON.stringify(prBody) });
      if (!r.ok) throw new Error(`create user progress: ${r.status} ${await r.text().catch(() => '')}`);
      logs.push({ step: 'Seed user progress', ok: true });
    } catch (e: any) {
      logs.push({ step: 'Seed user progress', ok: false, detail: String(e?.message || e) });
    }

    // 6) Create a User Acknowledgement row (optional)
    try {
      if (!batchId || !documentId) throw new Error('no batch/document id');
      const ackBody: any = {
        toba_name: 'Ack - jane.doe@contoso.com',
        toba_acknowledged: true,
        toba_ackdate: new Date().toISOString(),
        [DV_ATTRS.ackUserField]: 'jane.doe@contoso.com',
        [`${DV_ATTRS.ackBatchLookup}@odata.bind`]: `/${SETS.batchesSet}(${batchId})`,
        [`${DV_ATTRS.ackDocumentLookup}@odata.bind`]: `/${SETS.documentsSet}(${documentId})`
      };
      const r = await dvFetch(`/${SETS.userAcksSet}`, token, { method: 'POST', body: JSON.stringify(ackBody) });
      if (!r.ok) throw new Error(`create user acknowledgement: ${r.status} ${await r.text().catch(() => '')}`);
      logs.push({ step: 'Seed user acknowledgement', ok: true });
    } catch (e: any) {
      logs.push({ step: 'Seed user acknowledgement', ok: false, detail: String(e?.message || e) });
    }
  } catch (e: any) {
    logs.push({ step: 'Seed sample data failed', ok: false, detail: String(e?.message || e) });
  }
  return logs;
}

// Create -> Fetch -> Delete a minimal Batch record to verify write permissions, set names, and payloads.
export async function dvWriteTest(): Promise<ProvisionLog[]> {
  const logs: ProvisionLog[] = [];
  try {
    const token = await getDataverseToken();
    const SETS = await resolveEntitySets();
    try { logs.push({ step: 'Detect entity sets', ok: true, detail: JSON.stringify(SETS) }); } catch { logs.push({ step: 'Detect entity sets', ok: true }); }

    // Create minimal batch
    const name = `WriteTest ${new Date().toISOString()}`;
    const createRes = await dvFetch(`/${SETS.batchesSet}`, token, {
      method: 'POST',
      body: JSON.stringify({ toba_name: name })
    });
    if (!createRes.ok) {
      const t = await createRes.text().catch(() => '');
      logs.push({ step: 'Create test batch', ok: false, detail: `${createRes.status} ${t}` });
      return logs;
    }
    const loc = createRes.headers.get('OData-EntityId') || '';
    const m = loc.match(/[0-9a-fA-F-]{36}/);
    const id = m ? m[0] : undefined;
    logs.push({ step: 'Create test batch', ok: true, detail: id || 'no id' });
    if (!id) return logs; // cannot proceed

    // Fetch it back
    const fetchRes = await dvFetch(`/${SETS.batchesSet}(${id})?$select=toba_batchid,toba_name`, token, { method: 'GET' });
    if (!fetchRes.ok) {
      const t = await fetchRes.text().catch(() => '');
      logs.push({ step: 'Fetch test batch', ok: false, detail: `${fetchRes.status} ${t}` });
    } else {
      const j = await fetchRes.json().catch(() => ({}));
      logs.push({ step: 'Fetch test batch', ok: true, detail: JSON.stringify({ id, name: j?.toba_name }) });
    }

    // Delete it
    const delRes = await dvFetch(`/${SETS.batchesSet}(${id})`, token, { method: 'DELETE' });
    if (!delRes.ok) {
      const t = await delRes.text().catch(() => '');
      logs.push({ step: 'Delete test batch', ok: false, detail: `${delRes.status} ${t}` });
    } else {
      logs.push({ step: 'Delete test batch', ok: true, detail: id });
    }
  } catch (e: any) {
    logs.push({ step: 'DV write test failed', ok: false, detail: String(e?.message || e) });
  }
  return logs;
}
