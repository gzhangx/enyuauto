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
export const SP_ENYU_GENERAL_FOLDER = 'General';

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
  const data = await res.json();
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
    const item = await findRes.json();
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
  const created = await createRes.json();
  log(`Created file: ${MS_MAIN_EXCEL_FILE_NAME} (id: ${created.id})`);
  return { itemId: created.id as string, driveRoot };
}
