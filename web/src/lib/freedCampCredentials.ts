export const SECS_FILE_NAME = 'enyu_secs.xlsx';
export const FREED_CAMP_SHEET = 'freedCamp';
export const LOGIN_CFG_KEY = 'freedCamp_login_cfg';

export const MS_MAIN_EXCEL_FILE_NAME = 'enyustatus.xlsx';

export interface FreedCampCredentials {
  username: string;
  password: string;
}

async function getSecsFileId(msToken: string): Promise<string | null> {
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${SECS_FILE_NAME}`,
    { headers: { Authorization: `Bearer ${msToken}` } }
  );
  if (res.status === 404) return null;
  if (!res.ok) {
    const resJson = await res.json().catch(err=>err) as { error: { message: string; }} ;
    throw new Error(resJson.error?.message || res.statusText);
  }
  const resJson = (await res.json()) as  { id: string; };
  return resJson.id as string;
}

export async function readFreedCampCredentials(msToken: string): Promise<FreedCampCredentials | null> {
  const itemId = await getSecsFileId(msToken);
  if (!itemId) return null;

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/usedRange`,
    { headers: { Authorization: `Bearer ${msToken}` } }
  );
  if (!res.ok) return null;

  const jsonRet = (await res.json()) as { values: string[][]};
  const rows: string[][] = jsonRet.values ?? [];
  const row = rows.find(r => r[0] === LOGIN_CFG_KEY);
  if (!row) return null;
  return { username: row[1] ?? '', password: row[2] ?? '' };
}
