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

function encodeSharingUrl(url: string): string {
  const b64 = btoa(url).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
  return 'u!' + b64;
}

async function resolveSharePointSharingUrl(msToken: string, sharingUrl: string): Promise<{ id: string; name: string; parentReference?: { path?: string; driveId?: string } }> {
  const encoded = encodeSharingUrl(sharingUrl.trim());
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem?$select=id,name,parentReference`,
    { headers: { Authorization: `Bearer ${msToken}` } }
  );
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  return await res.json() as { id: string; name: string; parentReference?: { path?: string; driveId?: string } };
}

export interface ResolvedSharePointPath {
  driveId?: string;
  driveRootUrl?: string;
  path: string;
}

function normalizeSharePointWebUrl(value: string): { driveRootUrl: string; path: string } | null {
  try {
    const url = new URL(value);
    const pathname = decodeURIComponent(url.pathname);

    // Prefer explicit SharePoint query params that point to the actual file/folder
    const params = url.searchParams;
    const idParam = params.get('id');
    const viewPathParam = params.get('viewpath') || params.get('view') || params.get('newTargetListUrl');
    // If `id` is present it often contains the encoded path to the document library root
    if (idParam) {
      try {
        const decodedId = decodeURIComponent(idParam);
        // id often like "/sites/enyueditors/Shared Documents/General/..."
        const idMatch = decodedId.match(/^\/(sites|teams|personal)\/[^/]+\/(.*)/i);
        if (idMatch) {
          const siteRoot = decodedId.match(/^\/(?:sites|teams|personal)\/[^/]+/i)?.[0].slice(1); // remove leading /
          const relative = idMatch[2].replace(/^\/+/, '');
          const normalizedPath = relative.startsWith('Shared Documents/')
            ? relative.slice('Shared Documents/'.length)
            : relative;
          return {
            driveRootUrl: `https://graph.microsoft.com/v1.0/sites/${url.hostname}:/${siteRoot}:/drive`,
            path: normalizedPath,
          };
        }
      } catch {
        // fall through to other heuristics
      }
    }

    // If viewpath/newTargetListUrl contains the library path, prefer that
    if (viewPathParam) {
      try {
        const decoded = decodeURIComponent(viewPathParam);
        const siteRootMatch = decoded.match(/^\/(sites|teams|personal)\/[^/]+\//i);
        if (siteRootMatch) {
          const siteRoot = decoded.match(/^\/((?:sites|teams|personal)\/[^/]+)\//i)?.[1];
          const relativePath = decoded.slice(siteRootMatch[0].length).replace(/^\/+/, '');
          const normalizedPath = relativePath.startsWith('Shared Documents/')
            ? relativePath.slice('Shared Documents/'.length)
            : relativePath;
          return {
            driveRootUrl: `https://graph.microsoft.com/v1.0/sites/${url.hostname}:/${siteRoot}:/drive`,
            path: normalizedPath,
          };
        }
      } catch {
        // ignore
      }
    }

    // Fallback: parse pathname as before
    const siteRootMatch = pathname.match(/^\/(sites|teams|personal)\/[^/]+\//i);
    if (!siteRootMatch) return null;

    const siteRoot = pathname.match(/^\/((?:sites|teams|personal)\/[^/]+)\//i)?.[1];
    if (!siteRoot) return null;

    const relativePath = pathname.slice(siteRootMatch[0].length).replace(/^\/+/, '');
    const normalizedPath = relativePath.startsWith('Shared Documents/')
      ? relativePath.slice('Shared Documents/'.length)
      : relativePath;

    return {
      driveRootUrl: `https://graph.microsoft.com/v1.0/sites/${url.hostname}:/${siteRoot}:/drive`,
      path: normalizedPath,
    };
  } catch {
    return null;
  }
}

export async function resolveSharePointPath(rawValue: string, msToken?: string): Promise<ResolvedSharePointPath> {
  const originalValue = rawValue.trim();
  let value = originalValue;
  if (!value.startsWith('http://') && !value.startsWith('https://')) {
    return { path: value.replace(/^\/+/, '') };
  }
  if (value.startsWith('https://graph.microsoft.com')) {
    const driveRootMatch = value.match(/\/drive\/root:(\/.*?)(?:$|\/|\?)/);
    if (driveRootMatch) return { path: driveRootMatch[1].replace(/^\/+/, '') };
    const siteRootMatch = value.match(/\/sites\/[^/]+:[^/]+:.*?\/drive\/root:(\/.*?)(?:$|\/|\?)/);
    if (siteRootMatch) return { path: siteRootMatch[1].replace(/^\/+/, '') };
  }
  if (value.startsWith('https://')) {
    if (!msToken) {
      throw new Error(`Cannot resolve SharePoint sharing URL without msToken: ${rawValue}`);
    }

    // If the URL contains a sourcedoc (GUID) and a file name, try to resolve
    // the document by id first, then fall back to searching by filename.
    try {
      const parsed = new URL(value);
      const params = parsed.searchParams;
      const sourceDoc = params.get('sourcedoc');
      const fileName = params.get('file') ? decodeURIComponent(params.get('file') as string) : undefined;

      if (sourceDoc) {
        // Attempt to infer site from the path (e.g. /sites/enyueditors appears in pathname)
        const siteMatch = parsed.pathname.match(/\/sites\/([^/]+)/i)?.[1];
        const driveRootUrl = siteMatch
          ? `https://graph.microsoft.com/v1.0/sites/${parsed.hostname}:/sites/${siteMatch}:/drive`
          : undefined;

        // Try to fetch the item directly by id (sourcedoc might be the drive/item id)
        if (driveRootUrl) {
          try {
            const itemRes = await fetch(`${driveRootUrl}/items/${encodeURIComponent(sourceDoc)}`, {
              headers: { Authorization: `Bearer ${msToken}` },
            });
            if (itemRes.ok) {
              const item = await itemRes.json() as { id: string; name: string; parentReference?: { path?: string; driveId?: string } };
              const parentPath = item.parentReference?.path ?? '';
              const normalizedParentPath = parentPath.replace(/^\/drives?\/[^/]+\/root:/, '').replace(/^\/+/, '');
              return {
                driveId: item.parentReference?.driveId,
                driveRootUrl,
                path: `${normalizedParentPath}/${item.name}`.replace(/^\/+/, ''),
              };
            }
          } catch {
            // fall through to search
          }
        }

        // If direct id lookup failed, try searching by filename if available
        if (fileName && driveRootUrl) {
          try {
            const searchRes = await fetch(`${driveRootUrl}/root/search(q='${encodeURIComponent(fileName)}')`, {
              headers: { Authorization: `Bearer ${msToken}` },
            });
            if (searchRes.ok) {
              const searchData = await searchRes.json() as { value?: Array<any> };
              const match = (searchData.value || []).find((it: any) => it.name === fileName);
              if (match) {
                const parentPath = match.parentReference?.path ?? '';
                const normalizedParentPath = parentPath.replace(/^\/drives?\/[^/]+\/root:/, '').replace(/^\/+/, '');
                return {
                  driveId: match.parentReference?.driveId,
                  driveRootUrl,
                  path: `${normalizedParentPath}/${match.name}`.replace(/^\/+/, ''),
                };
              }
            }
          } catch {
            // fall through to generic sharing resolution below
          }
        }
      } else {
        const genericWeb = normalizeSharePointWebUrl(value);
        if (genericWeb) {
          return { driveRootUrl: genericWeb.driveRootUrl, path: genericWeb.path };
        }
      }

    } catch {
      // ignore URL parse errors and continue to generic sharing resolution
    }

    // Fallback to resolving a sharing URL via Graph if previous attempts didn't work
    const shareItem = await resolveSharePointSharingUrl(msToken, value);
    const shareItemPath = shareItem.parentReference?.path;
    if (!shareItemPath) {
      throw new Error(`Unable to resolve path from SharePoint sharing URL: ${rawValue}`);
    }
    const normalizedParentPath = shareItemPath.replace(/^\/drives?\/[^/]+\/root:/, '').replace(/^\/+/, '');
    return {
      driveId: shareItem.parentReference?.driveId,
      path: `${normalizedParentPath}/${shareItem.name}`.replace(/^\/+/, ''),
    };
  }
  return { path: value.replace(/^\/+/, '') };
}

export async function joinSharePointPath(msToken: string, basePath: string, relativePath: string): Promise<ResolvedSharePointPath> {
  const base = await resolveSharePointPath(basePath, msToken);
  const normalizedBase = base.path.replace(/[\/]+$/g, '');
  const normalizedRelative = relativePath.replace(/^[\/]+/g, '');
  return {
    driveId: base.driveId,
    driveRootUrl: base.driveRootUrl,
    path: `${normalizedBase}/${normalizedRelative}`,
  };
}


async function getDriveItemByPath(msToken: string, driveRoot: string, itemPath: string, driveId?: string, driveRootUrl?: string): Promise<{ id: string; name: string; webUrl: string }> {
  const path = itemPath.replace(/^\/+/, '');
  const url = driveId
    ? `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root:/${encodeURI(path)}`
    : driveRootUrl
      ? `${driveRootUrl}/root:/${encodeURI(path)}`
      : `${driveRoot}/root:/${encodeURI(path)}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${msToken}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  return await res.json() as { id: string; name: string; webUrl: string };
}

async function resolveDriveIdForRoot(msToken: string, driveRootUrl: string): Promise<string> {
  const res = await fetch(`${driveRootUrl}?$select=id`, {
    headers: { Authorization: `Bearer ${msToken}` },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || res.statusText);
  }
  const data = await res.json() as { id: string };
  return data.id;
}

interface WaitFileResultRes {
  webUrl: string;
  status: string;  //"failed",
  error: {
    code: string;  //"nameAlreadyExists",
    message: string; //"Name already exists",
    target: string; //;"01BDEZ3HBZS2ZJTK6F5RHYYUWYSG65AZTZ"
   }
}
async function waitForGraphCopyCompletion(location: string, authHeaders: Record<string, string>, log:DebugLog): Promise<string> {
  const start = Date.now();
  while (true) {
    const res = await fetch(location, { headers: authHeaders });
    if (res.status === 202) {
      log.doLog(`Waiting ${location}`)
      if (Date.now() - start > 20000) {
        log.doLog(`Waiting ${location} Tiomeout!!`)
        throw new Error('Timed out waiting for Graph copy operation to complete');
      }
      await new Promise((resolve) => setTimeout(resolve, 500));
      continue;
    }
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      log.doLog(`Waiting ${location} error! ${(body as any).error}`)
      //throw new Error((body as any)?.error?.message || res.statusText);
      return location;
    }
    const rr = await res.json() as WaitFileResultRes;
    log.doLog(`Waiting ${location} got res ${JSON.stringify(rr)}`);
    if (rr.status === 'failed' || rr.status === 'inProgress') {
      continue;
    }
    return location;
  }
}

interface DebugLog {
    doLog: (msg: string, critical?: boolean) => void;
}

export async function copySharePointFile(
  msToken: string,
  sourceFileUrl: string,
  destinationFolderUrl: string,
  log: DebugLog
): Promise<string> {
  const driveRoot = await resolveSiteGraphDriveRoot(msToken);
  const sourcePathResult = await resolveSharePointPath(sourceFileUrl, msToken);
  const destinationPathResult = await resolveSharePointPath(destinationFolderUrl, msToken);
  const destPath = destinationPathResult.path.replace(/\/+$/, '');

  const sourceItem = await getDriveItemByPath(
    msToken,
    driveRoot,
    sourcePathResult.path,
    sourcePathResult.driveId,
    sourcePathResult.driveRootUrl,
  );

  log.doLog(`Copy share point file ${sourceFileUrl} to ${destinationFolderUrl} resolved sourceItem ${sourceItem?.name}`);

  const destinationDriveId = destinationPathResult.driveId
    ?? (destinationPathResult.driveRootUrl ? await resolveDriveIdForRoot(msToken, destinationPathResult.driveRootUrl) : undefined);
  const parentReference = {
    driveId: destinationDriveId,
    path: destPath ? `/drive/root:/${destPath}` : '/drive/root:',
  };
  const sourceDriveRoot = sourcePathResult.driveId
    ? `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(sourcePathResult.driveId)}`
    : sourcePathResult.driveRootUrl ?? driveRoot;
  const copyRes = await fetch(`${sourceDriveRoot}/items/${encodeURIComponent(sourceItem.id)}/copy?@microsoft.graph.conflictBehavior=replace`, {
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
    log.doLog(`Copy share point file ${sourceFileUrl} to ${destinationFolderUrl} copy error`);
    const body = await copyRes.json().catch(() => ({}));
    throw new Error((body as any)?.error?.message || copyRes.statusText);
  }
  const location = copyRes.headers.get('Location');
  if (!location) {
    log.doLog(`Copy share point file ${sourceFileUrl} to ${destinationFolderUrl} no location`);
    throw new Error('Graph copy response did not include Location header');
  }
  log.doLog(`Copy share point file ${sourceFileUrl} to ${destinationFolderUrl} got location ${location}`);
  const result = await waitForGraphCopyCompletion(location, { Authorization: `Bearer ${msToken}` }, log);
  //result has this format:   {
//   "@odata.context": "https://acccnusa.sharepoint.com/sites/enyueditors/_api/v2.1/$metadata#drives('b!xwQ0UHWhIEWIjAbRaEc163pG3pJF4mNMiu8wVwNbWg1l_K7CJIsATIcOLrIKBWJY')/operations/$entity",
//   "id": "70f2a6d4-2472-4f48-b0c4-c9e8c9067f99",
//   "createdDateTime": "0001-01-01T00:00:00Z",
//   "lastActionDateTime": "0001-01-01T00:00:00Z",
//   "status": "failed",
//   "error": {
//     "code": "nameAlreadyExists",
//     "message": "Name already exists",
//     "target": "01BDEZ3HBZS2ZJTK6F5RHYYUWYSG65AZTZ"
//   }
  // }
  
  // After the copy operation completes, resolve the newly copied item's webUrl
  try {
    const itemPath = `${destPath}/${sourceItem.name}`.replace(/^\/+/, '');
    const newItem = await getDriveItemByPath(
      msToken,
      driveRoot,
      itemPath,
      destinationDriveId,
      destinationPathResult.driveRootUrl,
    );
    return newItem.webUrl;
  } catch (err) {
    // If we fail to resolve the copied item for any reason, fall back to returning
    // the operation location (so callers still get a non-empty string) and log.
    log.doLog(`copySharePointFile: failed to resolve copied item webUrl: ${(err as any)?.message || err}`);
    return result;
  }
}
