/**
 * Shared Microsoft Graph API configuration.
 * All Graph requests targeting the EnYu Editors SharePoint site should use
 * these constants so the root URL is maintained in one place.
 *
 * SharePoint site: https://acccnusa.sharepoint.com/sites/enyueditors
 * Library / folder: Shared Documents › General
 */

/** SharePoint folder browser URL (for reference/links only — not usable in fetch calls). */
export const SP_ENYU_DRIVE_ROOT =
  'https://acccnusa.sharepoint.com/sites/enyueditors/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2Fenyueditors%2FShared%20Documents%2FGeneral&viewid=8059c88a%2D1194%2D4065%2Db31a%2D3749f1c293f2';

/** Microsoft Graph API drive root derived from SP_ENYU_DRIVE_ROOT.
 *  Site: acccnusa.sharepoint.com/sites/enyueditors — default "Shared Documents" library.
 *  Use this constant in all fetch() / Graph API calls.
 */
export const SP_ENYU_GRAPH_DRIVE_ROOT =
  'https://graph.microsoft.com/v1.0/sites/acccnusa.sharepoint.com:/sites/enyueditors:/drive';

/** Sub-folder inside the document library where EnYu working files live. */
export const SP_ENYU_GENERAL_FOLDER = 'General/ENYU Co-workers Resources/Projects';

const SP_ENYU_SITE_GRAPH_URL = 'https://graph.microsoft.com/v1.0/sites/acccnusa.sharepoint.com:/sites/enyueditors';
let _resolvedDriveRoot: string | null = null;

/**
 * Resolves the proper Graph API drive root for the enyueditors SharePoint site.
 * Makes a single GET to resolve the site ID, then returns
 * `https://graph.microsoft.com/v1.0/sites/{id}/drive`.
 * Result is cached after the first call.
 */
export async function resolveSiteGraphDriveRoot(msToken: string): Promise<string> {
  if (_resolvedDriveRoot) return _resolvedDriveRoot;
  const res = await fetch(SP_ENYU_SITE_GRAPH_URL, {
    headers: { Authorization: `Bearer ${msToken}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  const data = await res.json() as { id: string };
  _resolvedDriveRoot = `https://graph.microsoft.com/v1.0/sites/${data.id}/drive`;
  return _resolvedDriveRoot;
}

/**
 * Finds the Excel workbook (MS_MAIN_EXCEL_FILE_NAME) in SP_ENYU_GENERAL_FOLDER.
 * If not found and `buildContent` is provided, creates the file with that content.
 * Returns `{ itemId, driveRoot }` for immediate use in subsequent Graph calls.
 */
export async function findOrCreateExcelFile(
  msToken: string,
  log: (s: string) => void = () => {},
  buildContent?: () => Uint8Array,
): Promise<{ itemId: string; driveRoot: string }> {
  const { MS_MAIN_EXCEL_FILE_NAME } = await import('./freedCampCredentials');
  const driveRoot = await resolveSiteGraphDriveRoot(msToken);
  const authHeaders = { Authorization: `Bearer ${msToken}` };

  const findRes = await fetch(
    `${driveRoot}/root:/${SP_ENYU_GENERAL_FOLDER}/${MS_MAIN_EXCEL_FILE_NAME}`,
    { headers: authHeaders },
  );
  if (findRes.ok) {
    const item = await findRes.json() as { id: string; };
    log(`Found existing file: ${MS_MAIN_EXCEL_FILE_NAME} (id: ${item.id})`);
    return { itemId: item.id as string, driveRoot };
  }
  if (findRes.status !== 404 || !buildContent) {
    const body = await findRes.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || findRes.statusText);
  }

  log(`File not found, creating ${MS_MAIN_EXCEL_FILE_NAME}…`);
  const content = buildContent();
  const createRes = await fetch(
    `${driveRoot}/root:/${SP_ENYU_GENERAL_FOLDER}/${MS_MAIN_EXCEL_FILE_NAME}:/content`,
    {
      method: 'PUT',
      headers: { ...authHeaders, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: content.buffer as ArrayBuffer,
    },
  );
  if (!createRes.ok) {
    const body = await createRes.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || createRes.statusText);
  }
  const created = await createRes.json() as { id: string; };
  log(`Created file: ${MS_MAIN_EXCEL_FILE_NAME} (id: ${created.id})`);
  return { itemId: created.id as string, driveRoot };
}

function normalizeSharePointPath(rawValue: string): string {
  let value = rawValue.trim();
  const qIndex = value.indexOf('?');
  if (qIndex >= 0) value = value.slice(0, qIndex);
  if (value.startsWith('https://graph.microsoft.com')) {
    const driveRootMatch = value.match(/\/drive\/root:(\/.*?)(?:$|\/|\?)/);
    if (driveRootMatch) return driveRootMatch[1].replace(/^\/+/, '');
    const siteRootMatch = value.match(/\/sites\/[^/]+:[^/]+:.*?\/drive\/root:(\/.*?)(?:$|\/|\?)/);
    if (siteRootMatch) return siteRootMatch[1].replace(/^\/+/, '');
  }
  if (value.startsWith('https://')) {
    const sitePathIndex = value.indexOf('/sites/enyueditors/');
    if (sitePathIndex >= 0) {
      const path = value.slice(sitePathIndex + '/sites/enyueditors/'.length);
      if (path.startsWith('Shared%20Documents/')) {
        return decodeURIComponent(path.slice('Shared%20Documents/'.length));
      }
      if (path.startsWith('Shared Documents/')) {
        return path.slice('Shared Documents/'.length);
      }
      return path.replace(/^\/+/, '');
    }
  }
  return value.replace(/^\/+/, '');
}

async function getDriveItemByPath(msToken: string, driveRoot: string, itemPath: string): Promise<{ id: string; name: string; webUrl: string }> {
  const path = itemPath.replace(/^\/+/, '');
  const res = await fetch(`${driveRoot}/root:/${encodeURI(path)}`, {
    headers: { Authorization: `Bearer ${msToken}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  return await res.json() as { id: string; name: string; webUrl: string };
}

async function waitForGraphCopyCompletion(location: string, authHeaders: Record<string, string>): Promise<{ webUrl: string }> {
  const start = Date.now();
  while (true) {
    const res = await fetch(location, { headers: authHeaders });
    if (res.status === 202) {
      if (Date.now() - start > 20000) {
        throw new Error('Timed out waiting for Graph copy operation to complete');
      }
      await new Promise((resolve) => setTimeout(resolve, 500));
      continue;
    }
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error((body as any)?.error?.message || res.statusText);
    }
    return await res.json() as { webUrl: string };
  }
}

export async function copySharePointFile(
  msToken: string,
  sourceFileUrl: string,
  destinationFolderUrl: string,
): Promise<string> {
  const driveRoot = await resolveSiteGraphDriveRoot(msToken);
  const sourcePath = normalizeSharePointPath(sourceFileUrl);
  const destPath = normalizeSharePointPath(destinationFolderUrl).replace(/\/+$/, '');

  const sourceItem = await getDriveItemByPath(msToken, driveRoot, sourcePath);
  const parentReference = {
    path: destPath ? `/drive/root:/${destPath}` : '/drive/root:',
  };
  const copyRes = await fetch(`${driveRoot}/items/${encodeURIComponent(sourceItem.id)}/copy`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${msToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      parentReference,
      name: sourceItem.name,
    }),
  });
  if (!copyRes.ok && copyRes.status !== 202) {
    const body = await copyRes.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || copyRes.statusText);
  }
  const location = copyRes.headers.get('Location');
  if (!location) {
    throw new Error('Graph copy response did not include Location header');
  }
  const result = await waitForGraphCopyCompletion(location, { Authorization: `Bearer ${msToken}` });
  return result.webUrl;
}
