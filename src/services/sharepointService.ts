/**
 * SharePoint Document Library Integration Service
 * Enables seamless document selection from SharePoint for batch creation
 */
import { info, warn, error as logError } from '../diagnostics/logger';

export interface SharePointSite {
  id: string;
  displayName: string;
  webUrl: string;
  description?: string;
}

export interface SharePointDocumentLibrary {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  webUrl: string;
  driveType: string;
}

export interface SharePointDocument {
  id: string;
  name: string;
  webUrl: string;
  size: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  file?: {
    mimeType: string;
    hashes?: {
      sha1Hash?: string;
    };
  };
  createdBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
  parentReference?: {
    driveId: string;
    driveType: string;
    path: string;
  };
}

/**
 * Upload a file to a SharePoint document library (Drive) root or optional folder path.
 * - Uses simple upload for files <= 4MB
 * - Uses Upload Session (chunked) for larger files
 * Returns the created DriveItem as SharePointDocument
 */
export const uploadFileToDrive = async (
  token: string,
  driveId: string,
  file: File | Blob,
  fileName?: string,
  folderPath?: string,
  onProgress?: (percent: number) => void,
  folderItemId?: string
): Promise<SharePointDocument> => {
  const name = fileName || (file instanceof File ? file.name : `upload-${Date.now()}`);
  const path = folderPath ? `${folderPath.replace(/^\/+|\/+$/g, '')}/${encodeURIComponent(name)}` : encodeURIComponent(name);

  // Simple upload path
  const simpleUpload = async () => {
    const target = folderItemId && folderItemId !== 'root'
      ? `items/${folderItemId}:/${encodeURIComponent(name)}:/content`
      : `root:/${path}:/content`;
    const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/${target}`, {
      method: 'PUT',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': (file as any).type || 'application/octet-stream' },
      body: file as any
    });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`Simple upload failed: ${res.status} ${res.statusText} — ${text}`);
    }
    const item = await res.json();
    return item as SharePointDocument;
  };

  // Chunked upload path
  const chunkedUpload = async () => {
    const target = folderItemId && folderItemId !== 'root'
      ? `items/${folderItemId}:/${encodeURIComponent(name)}:/createUploadSession`
      : `root:/${path}:/createUploadSession`;
    const sessionRes = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/${target}`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        item: { '@microsoft.graph.conflictBehavior': 'rename', name }
      })
    });
    if (!sessionRes.ok) {
      const text = await sessionRes.text().catch(() => '');
      throw new Error(`Create upload session failed: ${sessionRes.status} ${sessionRes.statusText} — ${text}`);
    }
    const session = await sessionRes.json();
    const uploadUrl = session.uploadUrl as string;

    const chunkSize = 5 * 1024 * 1024; // 5MB
    const total = (file as any).size ?? 0;
    let uploaded = 0;
    let start = 0;

    while (start < total) {
      const end = Math.min(start + chunkSize, total);
      const chunk = (file as any).slice(start, end);
      const contentRange = `bytes ${start}-${end - 1}/${total}`;
      const res = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(end - start),
          'Content-Range': contentRange
        },
        body: chunk
      });
      if (!res.ok && res.status !== 202) {
        const text = await res.text().catch(() => '');
        throw new Error(`Chunk upload failed: ${res.status} ${res.statusText} — ${text}`);
      }
      uploaded = end;
      if (onProgress) onProgress(Math.round((uploaded / total) * 100));

      // If completed, Graph returns the DriveItem
      if (res.status === 201 || res.status === 200) {
        const item = await res.json();
        return item as SharePointDocument;
      }
      start = end;
    }
    // Fallback fetch item (should not reach if above returned)
    throw new Error('Upload session did not return a completed item.');
  };

  if ((file as any).size && (file as any).size <= 4 * 1024 * 1024) {
    return await simpleUpload();
  }
  return await chunkedUpload();
};

/**
 * Gets SharePoint sites accessible to the user
 */
export const getSharePointSites = async (token: string): Promise<SharePointSite[]> => {
  try {
    const response = await fetch(
      'https://graph.microsoft.com/v1.0/sites?search=*&$select=id,displayName,webUrl,description&$top=50',
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!response.ok) {
      throw new Error(`SharePoint API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const sites = data.value as SharePointSite[];

    info('sharepointService: fetched sites', { count: sites.length });
    return sites;
  } catch (e) {
    logError('sharepointService: failed to fetch sites', e);
    throw e;
  }
};

/**
 * Gets document libraries for a specific SharePoint site
 */
export const getDocumentLibraries = async (token: string, siteId: string): Promise<SharePointDocumentLibrary[]> => {
  const sitePath = siteId; // Graph site-id path segments typically include commas and should not be URL-encoded
  // Helper to map drives payload into our shape, using drive.name as displayName
  const mapDrives = (arr: any[]): SharePointDocumentLibrary[] =>
    (arr || [])
      .filter((drive: any) => drive.driveType === 'documentLibrary')
      .map((drive: any) => ({
        id: drive.id,
        name: drive.name,
        displayName: drive.name, // drive doesn't have displayName; mirror name
        description: '',
        webUrl: drive.webUrl,
        driveType: drive.driveType
      }));

  // Fallback via Lists API
  const fetchViaLists = async (): Promise<SharePointDocumentLibrary[]> => {
    const listsUrl = `https://graph.microsoft.com/v1.0/sites/${sitePath}/lists?$filter=drive ne null&$select=id,displayName,webUrl&$expand=drive($select=id,driveType,name,webUrl)`;
    const listsRes = await fetch(listsUrl, { headers: { Authorization: `Bearer ${token}` } });
    if (!listsRes.ok) {
      const text = await listsRes.text().catch(() => '');
      throw new Error(`Lists fallback failed: ${listsRes.status} ${listsRes.statusText}${text ? ' — ' + text : ''}`);
    }
    const listsData = await listsRes.json();
    const libraries = (listsData.value || [])
      .filter((l: any) => l.drive && (l.drive.driveType === 'documentLibrary' || l.drive.id))
      .map((l: any) => ({
        id: l.drive?.id || l.id,
        name: l.drive?.name || l.displayName,
        displayName: l.displayName || l.drive?.name,
        description: '',
        webUrl: l.webUrl,
        driveType: (l.drive?.driveType || 'documentLibrary')
      })) as SharePointDocumentLibrary[];
    info('sharepointService: fetched document libraries (lists fallback)', { siteId, count: libraries.length });
    return libraries;
  };

  // Primary via Drives API
  try {
    const url = `https://graph.microsoft.com/v1.0/sites/${sitePath}/drives?$select=id,name,webUrl,driveType&$top=200`;
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!response.ok) {
      const text = await response.text().catch(() => '');
      // Attempt fallback instead of failing immediately
      info('sharepointService: drives listing not ok, attempting lists fallback', { status: response.status });
      try { return await fetchViaLists(); } catch (fallbackErr) {
        throw new Error(`Drives listing failed: ${response.status} — ${text || response.statusText}. Fallback error: ${(fallbackErr as Error).message}`);
      }
    }
    const data = await response.json();
    let libraries = mapDrives(data.value || []);
    if (libraries.length === 0) {
      info('sharepointService: drives empty, attempting lists fallback');
      libraries = await fetchViaLists();
    }
    info('sharepointService: fetched document libraries (drives)', { siteId, count: libraries.length });
    return libraries;
  } catch (e) {
    logError('sharepointService: failed to fetch document libraries', e);
    // As a last resort, try lists once if not already tried inside
    try { return await fetchViaLists(); } catch { throw e; }
  }
};

/**
 * Gets documents from a specific document library
 */
export const getDocuments = async (
  token: string, 
  driveId: string, 
  folderId: string = 'root',
  searchQuery?: string
): Promise<SharePointDocument[]> => {
  try {
    let url = folderId === 'root'
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children`;
    url += '?$select=id,name,webUrl,size,createdDateTime,lastModifiedDateTime,file,createdBy,lastModifiedBy,parentReference';
    url += '&$expand=thumbnails&$top=200';

    if (searchQuery) {
      url = `https://graph.microsoft.com/v1.0/drives/${driveId}/search(q='${encodeURIComponent(searchQuery)}')`;
      url += '?$select=id,name,webUrl,size,createdDateTime,lastModifiedDateTime,file,createdBy,lastModifiedBy,parentReference';
    }

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
      throw new Error(`SharePoint API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    // Filter to files only (exclude folders)
    const documents = data.value.filter((item: any) => item.file) as SharePointDocument[];

    // Filter to common document types
    const supportedTypes = [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'text/plain',
      'text/html'
    ];

    const filteredDocs = documents.filter(doc => 
      doc.file && supportedTypes.includes(doc.file.mimeType)
    );

    info('sharepointService: fetched documents', { driveId, folderId, count: filteredDocs.length, searchQuery });
    return filteredDocs;
  } catch (e) {
    logError('sharepointService: failed to fetch documents', e);
    throw e;
  }
};

/**
 * Gets document content/download URL for preview
 */
export const getDocumentDownloadUrl = async (token: string, driveId: string, itemId: string): Promise<string> => {
  try {
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}?$select=@microsoft.graph.downloadUrl`,
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!response.ok) {
      throw new Error(`SharePoint API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    return data['@microsoft.graph.downloadUrl'];
  } catch (e) {
    logError('sharepointService: failed to get download URL', e);
    throw e;
  }
};

/**
 * Gets document metadata including version history
 */
export const getDocumentMetadata = async (token: string, driveId: string, itemId: string): Promise<any> => {
  try {
    const [itemResponse, versionsResponse] = await Promise.all([
      fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`, {
        headers: { Authorization: `Bearer ${token}` }
      }),
      fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/versions`, {
        headers: { Authorization: `Bearer ${token}` }
      })
    ]);

    const item = itemResponse.ok ? await itemResponse.json() : {};
    const versions = versionsResponse.ok ? (await versionsResponse.json()).value : [];

    return {
      ...item,
      versions: versions
    };
  } catch (e) {
    logError('sharepointService: failed to get document metadata', e);
    throw e;
  }
};

/**
 * Creates a webhook subscription for document library changes
 */
export const createDocumentLibrarySubscription = async (
  token: string, 
  driveId: string, 
  notificationUrl: string
): Promise<any> => {
  try {
    const subscription = {
      changeType: 'created,updated',
      notificationUrl: notificationUrl,
      resource: `/drives/${driveId}/root`,
      expirationDateTime: new Date(Date.now() + 4230 * 60 * 1000).toISOString(), // ~3 days
      clientState: 'sunbeth-portal-webhook'
    };

    const response = await fetch('https://graph.microsoft.com/v1.0/subscriptions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(subscription)
    });

    if (!response.ok) {
      throw new Error(`SharePoint subscription error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    info('sharepointService: created document library subscription', { driveId, subscriptionId: data.id });
    return data;
  } catch (e) {
    logError('sharepointService: failed to create subscription', e);
    throw e;
  }
};

/**
 * List children (files and folders) of a folder in a drive. Useful for folder picking.
 */
export const getFolderItems = async (
  token: string,
  driveId: string,
  folderId: string = 'root'
): Promise<any[]> => {
  try {
    const base = folderId === 'root'
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children`;
    const url = `${base}?$select=id,name,folder,file,webUrl,size,createdDateTime,lastModifiedDateTime,parentReference&$top=200`;
    const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!response.ok) {
      throw new Error(`SharePoint API error: ${response.status} ${response.statusText}`);
    }
    const data = await response.json();
    return data.value || [];
  } catch (e) {
    logError('sharepointService: failed to fetch folder items', e);
    throw e;
  }
};