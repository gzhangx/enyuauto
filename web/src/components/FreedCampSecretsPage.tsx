import { useState, useEffect } from 'react';
import { useAuth } from '../contexts/AuthContext';
import { FREED_CAMP_SHEET, LOGIN_CFG_KEY, type FreedCampCredentials } from '../lib/freedCampCredentials';

const SECS_FILE_NAME = 'enyu_secs.xlsx';
const EXCEL_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// ---- Minimal XLSX builder (same CRC32/ZIP helpers as TransfersPage) ----
const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();
function crc32(data: Uint8Array) {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = CRC32_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function w16(v: DataView, o: number, n: number) { v.setUint16(o, n, true); }
function w32(v: DataView, o: number, n: number) { v.setUint32(o, n, true); }
function concat(...parts: Uint8Array[]) {
  const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
  let pos = 0; for (const p of parts) { out.set(p, pos); pos += p.length; }
  return out;
}
function buildZip(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const enc = new TextEncoder();
  const local: Uint8Array[] = [], cd: Uint8Array[] = [];
  let off = 0;
  for (const f of files) {
    const nb = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
    const lh = new DataView(new ArrayBuffer(30));
    w32(lh, 0, 0x04034b50); w16(lh, 4, 20); w16(lh, 6, 0); w16(lh, 8, 0);
    w16(lh, 10, 0); w16(lh, 12, 0); w32(lh, 14, crc); w32(lh, 18, sz); w32(lh, 22, sz);
    w16(lh, 26, nb.length); w16(lh, 28, 0);
    local.push(new Uint8Array(lh.buffer), nb, f.data);
    const ch = new DataView(new ArrayBuffer(46));
    w32(ch, 0, 0x02014b50); w16(ch, 4, 20); w16(ch, 6, 20); w16(ch, 8, 0);
    w16(ch, 10, 0); w16(ch, 12, 0); w16(ch, 14, 0); w32(ch, 16, crc); w32(ch, 20, sz);
    w32(ch, 24, sz); w16(ch, 28, nb.length); w16(ch, 30, 0); w16(ch, 32, 0);
    w16(ch, 34, 0); w16(ch, 36, 0); w32(ch, 38, 0); w32(ch, 42, off);
    cd.push(new Uint8Array(ch.buffer), nb);
    off += 30 + nb.length + sz;
  }
  const cdBytes = concat(...cd);
  const eocd = new DataView(new ArrayBuffer(22));
  w32(eocd, 0, 0x06054b50); w16(eocd, 4, 0); w16(eocd, 6, 0);
  w16(eocd, 8, files.length); w16(eocd, 10, files.length);
  w32(eocd, 12, cdBytes.length); w32(eocd, 16, off); w16(eocd, 20, 0);
  return concat(...local, cdBytes, new Uint8Array(eocd.buffer));
}
function buildMinimalXlsx(sheetName: string): Uint8Array {
  const e = new TextEncoder();
  const safeName = sheetName.replace(/['"<>&]/g, '_');
  return buildZip([
    { name: '[Content_Types].xml', data: e.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>') },
    { name: '_rels/.rels', data: e.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
    { name: 'xl/workbook.xml', data: e.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${safeName}" sheetId="1" r:id="rId1"/></sheets></workbook>`) },
    { name: 'xl/_rels/workbook.xml.rels', data: e.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') },
    { name: 'xl/worksheets/sheet1.xml', data: e.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>') },
    { name: 'xl/styles.xml', data: e.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>') },
  ]);
}

// ---- Graph helpers ----
async function findOrCreateSecsFile(msToken: string): Promise<string> {
  const auth = { Authorization: `Bearer ${msToken}` };
  const findRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${SECS_FILE_NAME}`, { headers: auth });
  if (findRes.ok) return (await findRes.json()).id as string;
  if (findRes.status !== 404) throw new Error((await findRes.json().catch(() => ({}))).error?.message || findRes.statusText);

  const bytes = buildMinimalXlsx(FREED_CAMP_SHEET);
  const createRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${SECS_FILE_NAME}:/content`,
    { method: 'PUT', headers: { ...auth, 'Content-Type': EXCEL_CONTENT_TYPE }, body: bytes.buffer as ArrayBuffer }
  );
  if (!createRes.ok) throw new Error((await createRes.json().catch(() => ({}))).error?.message || createRes.statusText);
  return (await createRes.json()).id as string;
}

async function readCredentials(msToken: string): Promise<FreedCampCredentials | null> {
  const itemId = await findOrCreateSecsFile(msToken);
  const auth = { Authorization: `Bearer ${msToken}` };

  // Ensure the sheet exists
  const wsListRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets`, { headers: auth });
  if (wsListRes.ok) {
    const sheets: { name: string }[] = (await wsListRes.json()).value;
    if (!sheets.find(s => s.name === FREED_CAMP_SHEET)) {
      await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets/add`, {
        method: 'POST',
        headers: { ...auth, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: FREED_CAMP_SHEET }),
      });
    }
  }

  const rangeRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/usedRange`,
    { headers: auth }
  );
  if (!rangeRes.ok) return null;
  const rows: string[][] = (await rangeRes.json()).values ?? [];
  const row = rows.find(r => r[0] === LOGIN_CFG_KEY) || [];  
  const enyuTokenRow = rows.find(r => r[0] === 'enyu_wp_token') || [];
  const ret: FreedCampCredentials = {
    username: row[1] ?? '', password: row[2] ?? '',
    enyu_wp_token: enyuTokenRow[1] ?? '',
  };
  return ret;
}

async function saveCredentials(msToken: string, save: FreedCampCredentials): Promise<void> {
  const itemId = await findOrCreateSecsFile(msToken);
  const auth = { Authorization: `Bearer ${msToken}`, 'Content-Type': 'application/json' };

  // Read existing rows so we can find/replace the login_cfg row
  const rangeRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/usedRange`,
    { headers: { Authorization: `Bearer ${msToken}` } }
  );
  let rows: string[][] = [];
  if (rangeRes.ok) rows = (await rangeRes.json()).values ?? [];

  const rowIdx = rows.findIndex(r => r[0] === LOGIN_CFG_KEY);
  const newRow = [LOGIN_CFG_KEY, save.username, save.password];

  if (rowIdx === -1) {
    // Append after last row (1-indexed)
    const writeRow = rows.length + 1;
    const address = `A${writeRow}:C${writeRow}`;
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/range(address='${address}')`,
      { method: 'PATCH', headers: auth, body: JSON.stringify({ values: [newRow] }) }
    );
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error?.message || res.statusText);
  } else {
    // Update in-place (spreadsheet row = rowIdx + 1, 1-indexed)
    const writeRow = rowIdx + 1;
    const address = `A${writeRow}:C${writeRow}`;
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/workbook/worksheets('${encodeURIComponent(FREED_CAMP_SHEET)}')/range(address='${address}')`,
      { method: 'PATCH', headers: auth, body: JSON.stringify({ values: [newRow] }) }
    );
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error?.message || res.statusText);
  }
}

// ---- Component ----
const btnStyle = (color: string, disabled = false): React.CSSProperties => ({
  padding: '0.5rem 1.25rem',
  backgroundColor: disabled ? '#aaa' : color,
  color: 'white',
  border: 'none',
  borderRadius: '4px',
  fontSize: '0.95rem',
  cursor: disabled ? 'not-allowed' : 'pointer',
});

const inputStyle: React.CSSProperties = {
  padding: '0.45rem 0.75rem',
  fontSize: '0.95rem',
  border: '1px solid #ccc',
  borderRadius: '4px',
  width: '320px',
};

export const FreedCampSecretsPage = () => {
  const { msToken, msAccount, msLoginRedirect, freedCampCredentials, setFreedCampCredentials } = useAuth();
  const [username, setUsername] = useState(freedCampCredentials?.username ?? '');
  const [password, setPassword] = useState(freedCampCredentials?.password ?? '');
  const [enyu_wp_token, setenyu_wp_token] = useState(freedCampCredentials?.enyu_wp_token ?? '');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState('');
  const [error, setError] = useState('');

  // Keep local fields in sync when context changes (e.g. loaded from another tab)
  useEffect(() => {
    if (freedCampCredentials) {
      setUsername(freedCampCredentials.username);
      setPassword(freedCampCredentials.password);
      setenyu_wp_token(freedCampCredentials.enyu_wp_token ?? '');
    }
  }, [freedCampCredentials]);

  const handleLoad = async () => {
    if (!msToken) return;
    setLoading(true);
    setStatus('');
    setError('');
    try {
      const creds = await readCredentials(msToken);
      if (creds) {
        setUsername(creds.username);
        setPassword(creds.password);
        setenyu_wp_token(creds.enyu_wp_token);
        setFreedCampCredentials(creds);
        setStatus('Loaded successfully.');
      } else {
        setStatus(`No "${LOGIN_CFG_KEY}" row found in ${SECS_FILE_NAME} → ${FREED_CAMP_SHEET}. Fill in and Save.`);
      }
    } catch (e: any) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  };

  const handleSave = async () => {
    if (!msToken) return;
    setSaving(true);
    setStatus('');
    setError('');
    try {
      const saveObj = { username, password, enyu_wp_token };
      await saveCredentials(msToken, saveObj);
      setFreedCampCredentials(saveObj);
      setStatus('Saved successfully.');
    } catch (e: any) {
      setError(e.message || String(e));
    } finally {
      setSaving(false);
    }
  };

  if (!msAccount || !msToken) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '40vh', gap: '1rem' }}>
        <p style={{ color: '#555' }}>Sign in with Microsoft to manage FreedCamp credentials.</p>
        <button onClick={msLoginRedirect} style={btnStyle('#0078d4')}>Sign in with Microsoft</button>
      </div>
    );
  }

  return (
    <div style={{ padding: '1.5rem', maxWidth: '600px', margin: '0 auto' }}>
      <h2 style={{ marginBottom: '0.25rem' }}>FreedCamp Login Info</h2>
      <p style={{ color: '#666', fontSize: '0.9rem', marginBottom: '1.5rem' }}>
        Stored in <strong>{SECS_FILE_NAME}</strong> → sheet <strong>{FREED_CAMP_SHEET}</strong> on your OneDrive root.
        The file and sheet are created automatically if they do not exist.
      </p>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '0.85rem', marginBottom: '1.25rem' }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem', fontSize: '0.9rem' }}>
          Username
          <input
            type="text"
            value={username}
            onChange={e => setUsername(e.target.value)}
            autoComplete="username"
            style={inputStyle}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem', fontSize: '0.9rem' }}>
          Password
          <input
            type="password"
            value={password}
            onChange={e => setPassword(e.target.value)}
            autoComplete="current-password"
            style={inputStyle}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '0.3rem', fontSize: '0.9rem' }}>
          EnYu WP Token
          <input
            type="text"
            value={enyu_wp_token}
            onChange={e => setenyu_wp_token(e.target.value)}
            style={inputStyle}
          />
        </label>
      </div>

      <div style={{ display: 'flex', gap: '0.75rem', flexWrap: 'wrap', alignItems: 'center' }}>
        <button onClick={handleLoad} disabled={loading || saving} style={btnStyle('#0078d4', loading || saving)}>
          {loading ? 'Loading…' : 'Load from OneDrive'}
        </button>
        <button onClick={handleSave} disabled={saving || loading} style={btnStyle('#107c41', saving || loading)}>
          {saving ? 'Saving…' : 'Save to OneDrive'}
        </button>
      </div>

      {status && <p style={{ marginTop: '1rem', color: '#107c41', fontSize: '0.9rem' }}>{status}</p>}
      {error && <p style={{ marginTop: '1rem', color: '#dc3545', fontSize: '0.9rem' }}>{error}</p>}

      {freedCampCredentials && (
        <p style={{ marginTop: '1.25rem', fontSize: '0.85rem', color: '#555' }}>
          Active in memory: <strong>{freedCampCredentials.username}</strong> (password set: {freedCampCredentials.password ? 'yes' : 'no'})
        </p>
      )}
    </div>
  );
};
