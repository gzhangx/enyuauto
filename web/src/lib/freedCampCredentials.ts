export const SECS_FILE_NAME = 'enyu_secs.xlsx';
export const FREED_CAMP_SHEET = 'freedCamp';
export const LOGIN_CFG_KEY = 'freedCamp_login_cfg';

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
  if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error?.message || res.statusText);
  return (await res.json()).id as string;
}

export async function readFreedCampCredentials(msToken: string): Promise<FreedCampCredentials | null> {
  const itemId = await getSecsFileId(msToken);
  if (!itemId) return null;

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/usedRange`,
    { headers: { Authorization: `Bearer ${msToken}` } }
  );
  if (!res.ok) return null;

  const rows: string[][] = (await res.json()).values ?? [];
  const row = rows.find(r => r[0] === LOGIN_CFG_KEY);
  if (!row) return null;
  return { username: row[1] ?? '', password: row[2] ?? '' };
}
