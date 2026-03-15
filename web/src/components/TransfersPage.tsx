import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';
import { MS_MAIN_EXCEL_FILE_NAME } from '../lib/freedCampCredentials';

const GOOGLE_SHEET_ID = '1zSPJudO0DERn74xV2auIXeNbJxh1apO0tjzB4IrTeQk';

const EXCEL_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// ---- Minimal ZIP / XLSX builder ----

const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) crc = CRC32_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function w16(v: DataView, o: number, n: number) { v.setUint16(o, n, true); }
function w32(v: DataView, o: number, n: number) { v.setUint32(o, n, true); }

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const p of parts) { out.set(p, pos); pos += p.length; }
  return out;
}

function buildZip(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const enc = new TextEncoder();
  const localChunks: Uint8Array[] = [];
  const cdChunks: Uint8Array[] = [];
  let localOffset = 0;

  for (const file of files) {
    const nameBytes = enc.encode(file.name);
    const crc = crc32(file.data);
    const size = file.data.length;

    const lh = new DataView(new ArrayBuffer(30));
    w32(lh, 0, 0x04034b50); w16(lh, 4, 20);  w16(lh, 6, 0);  w16(lh, 8, 0);
    w16(lh, 10, 0);          w16(lh, 12, 0);  w32(lh, 14, crc); w32(lh, 18, size);
    w32(lh, 22, size);       w16(lh, 26, nameBytes.length); w16(lh, 28, 0);
    localChunks.push(new Uint8Array(lh.buffer), nameBytes, file.data);

    const cd = new DataView(new ArrayBuffer(46));
    w32(cd, 0, 0x02014b50); w16(cd, 4, 20); w16(cd, 6, 20); w16(cd, 8, 0);
    w16(cd, 10, 0); w16(cd, 12, 0); w16(cd, 14, 0); w32(cd, 16, crc);
    w32(cd, 20, size); w32(cd, 24, size); w16(cd, 28, nameBytes.length);
    w16(cd, 30, 0); w16(cd, 32, 0); w16(cd, 34, 0); w16(cd, 36, 0);
    w32(cd, 38, 0); w32(cd, 42, localOffset);
    cdChunks.push(new Uint8Array(cd.buffer), nameBytes);

    localOffset += 30 + nameBytes.length + size;
  }

  const cdBytes = concatBytes(cdChunks);
  const eocd = new DataView(new ArrayBuffer(22));
  w32(eocd, 0, 0x06054b50); w16(eocd, 4, 0); w16(eocd, 6, 0);
  w16(eocd, 8, files.length); w16(eocd, 10, files.length);
  w32(eocd, 12, cdBytes.length); w32(eocd, 16, localOffset); w16(eocd, 20, 0);

  return concatBytes([...localChunks, cdBytes, new Uint8Array(eocd.buffer)]);
}

function buildMinimalXlsx(): Uint8Array {
  const e = new TextEncoder();
  return buildZip([
    {
      name: '[Content_Types].xml', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'
      )
    },
    {
      name: '_rels/.rels', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
      )
    },
    {
      name: 'xl/workbook.xml', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
        '</workbook>'
      )
    },
    {
      name: 'xl/_rels/workbook.xml.rels', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'
      )
    },
    {
      name: 'xl/worksheets/sheet1.xml', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
      )
    },
    {
      name: 'xl/styles.xml', data: e.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
        '</styleSheet>'
      )
    },
  ]);
}

// ---- Column number to Excel letter (1=A, 26=Z, 27=AA …) ----
function toColLetter(col: number): string {
  let result = '';
  let n = col;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

// ---- Graph helpers ----
async function findOrCreateExcelFile(msToken: string, log: (s: string) => void): Promise<string> {
  const authHeaders = { Authorization: `Bearer ${msToken}` };

  const findRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${MS_MAIN_EXCEL_FILE_NAME}`,
    { headers: authHeaders }
  );
  if (findRes.ok) {
    const item = await findRes.json();
    log(`Found existing file: ${MS_MAIN_EXCEL_FILE_NAME} (id: ${item.id})`);
    return item.id as string;
  }
  if (findRes.status !== 404) {
    const body = await findRes.json().catch(() => ({}));
    throw new Error(body?.error?.message || findRes.statusText);
  }

  log(`File not found, creating ${MS_MAIN_EXCEL_FILE_NAME}…`);
  const minimalXlsx = buildMinimalXlsx();
  const createRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root:/${MS_MAIN_EXCEL_FILE_NAME}:/content`,
    { method: 'PUT', headers: { ...authHeaders, 'Content-Type': EXCEL_CONTENT_TYPE }, body: minimalXlsx.buffer as ArrayBuffer }
  );
  if (!createRes.ok) {
    const body = await createRes.json().catch(() => ({}));
    throw new Error(body?.error?.message || createRes.statusText);
  }
  const created = await createRes.json();
  log(`Created file: ${MS_MAIN_EXCEL_FILE_NAME} (id: ${created.id})`);
  return created.id as string;
}

async function copyGoogleSheetToOneDrive(
  googleToken: string,
  msToken: string,
  log: (s: string) => void,
  setProgress: (s: string) => void,
): Promise<void> {
  // 1. Fetch sheet list
  log('Fetching Google Sheet metadata…');
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}?fields=sheets.properties`,
    { headers: { Authorization: `Bearer ${googleToken}` } }
  );
  if (!metaRes.ok) {
    const body = await metaRes.json().catch(() => ({}));
    throw new Error(body?.error?.message || metaRes.statusText);
  }
  const meta = await metaRes.json();
  const sheetTitles: string[] = (meta.sheets as any[]).map(s => s.properties.title as string);
  log(`Found ${sheetTitles.length} sheet(s): ${sheetTitles.join(', ')}`);

  // 2. Find or create Excel file
  const itemId = await findOrCreateExcelFile(msToken, log);
  const baseUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`;
  const jsonHeaders: Record<string, string> = {
    Authorization: `Bearer ${msToken}`,
    'Content-Type': 'application/json',
  };

  // 3. Create persistent workbook session
  log('Creating workbook session…');
  let sessionId: string | null = null;
  const sessionRes = await fetch(`${baseUrl}/workbook/createSession`, {
    method: 'POST',
    headers: jsonHeaders,
    body: JSON.stringify({ persistChanges: true }),
  });
  if (sessionRes.ok) {
    sessionId = (await sessionRes.json()).id ?? null;
    log(`Session created: ${sessionId}`);
  } else {
    log('Warning: could not create workbook session, proceeding without one');
  }

  const wsHeaders: Record<string, string> = { ...jsonHeaders };
  if (sessionId) wsHeaders['workbook-session-id'] = sessionId;

  const closeSession = () => {
    if (!sessionId) return;
    fetch(`${baseUrl}/workbook/closeSession`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${msToken}`, 'workbook-session-id': sessionId! },
    }).catch(() => {});
  };

  try {
    // 4. List existing worksheets in Excel
    const wsListRes = await fetch(`${baseUrl}/workbook/worksheets`, { headers: wsHeaders });
    if (!wsListRes.ok) {
      const body = await wsListRes.json().catch(() => ({}));
      throw new Error(body?.error?.message || wsListRes.statusText);
    }
    const existingNames = new Set<string>((await wsListRes.json()).value.map((ws: any) => ws.name as string));

    // 5. Process each Google sheet
    for (const title of sheetTitles) {
      setProgress(`Processing: ${title}`);

      // Read data from Google Sheets
      log(`Reading "${title}" from Google Sheets…`);
      const dataRes = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${GOOGLE_SHEET_ID}/values/${encodeURIComponent(title)}`,
        { headers: { Authorization: `Bearer ${googleToken}` } }
      );
      const rows: any[][] = dataRes.ok ? ((await dataRes.json()).values ?? []) : [];
      log(`  ${rows.length} row(s) read`);

      // Ensure worksheet exists in Excel
      if (!existingNames.has(title)) {
        log(`  Adding worksheet "${title}"…`);
        const addRes = await fetch(`${baseUrl}/workbook/worksheets/add`, {
          method: 'POST',
          headers: wsHeaders,
          body: JSON.stringify({ name: title }),
        });
        if (!addRes.ok) {
          const body = await addRes.json().catch(() => ({}));
          throw new Error(`Failed to add worksheet "${title}": ${body?.error?.message || addRes.statusText}`);
        }
        existingNames.add(title);
      }

      if (rows.length === 0) {
        log(`  Sheet is empty, skipping write`);
        continue;
      }

      // Pad rows to uniform column width
      const numCols = Math.max(...rows.map(r => r.length));
      const numRows = rows.length;
      const paddedRows = rows.map(r => {
        const row = r.map(c => (c === null || c === undefined ? '' : String(c)));
        while (row.length < numCols) row.push('');
        return row;
      });

      // Clear existing contents
      await fetch(
        `${baseUrl}/workbook/worksheets('${encodeURIComponent(title)}')/usedRange/clear`,
        { method: 'POST', headers: wsHeaders, body: JSON.stringify({ applyTo: 'Contents' }) }
      ).catch(() => { /* non-fatal if sheet was blank */ });

      // Write data
      const address = `A1:${toColLetter(numCols)}${numRows}`;
      log(`  Writing to ${title}!${address}…`);
      const writeRes = await fetch(
        `${baseUrl}/workbook/worksheets('${encodeURIComponent(title)}')/range(address='${address}')`,
        { method: 'PATCH', headers: wsHeaders, body: JSON.stringify({ values: paddedRows }) }
      );
      if (!writeRes.ok) {
        const body = await writeRes.json().catch(() => ({}));
        throw new Error(`Failed to write "${title}": ${body?.error?.message || writeRes.statusText}`);
      }
      log(`  ✓ ${title} done`);
    }
  } finally {
    closeSession();
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

export const TransfersPage = () => {
  const { token, msToken, msAccount, msLoginRedirect } = useAuth();
  const [status, setStatus] = useState<'idle' | 'running' | 'done' | 'error'>('idle');
  const [logs, setLogs] = useState<string[]>([]);
  const [progress, setProgress] = useState('');

  const addLog = (msg: string) => setLogs(prev => [...prev, msg]);

  const handleCopy = async () => {
    if (!token || !msToken) return;
    setStatus('running');
    setLogs([]);
    setProgress('');
    try {
      await copyGoogleSheetToOneDrive(token, msToken, addLog, setProgress);
      addLog('✓ Transfer complete!');
      setStatus('done');
    } catch (err: any) {
      addLog(`✗ Error: ${err.message || String(err)}`);
      setStatus('error');
    } finally {
      setProgress('');
    }
  };

  const running = status === 'running';

  return (
    <div style={{ padding: '1.5rem', maxWidth: '800px', margin: '0 auto' }}>
      <h2 style={{ marginBottom: '0.5rem' }}>Transfers</h2>
      <p style={{ color: '#555', marginBottom: '1.5rem', fontSize: '0.95rem' }}>
        Copies all sheets from the main Google Sheet into <strong>{MS_MAIN_EXCEL_FILE_NAME}</strong> in your OneDrive root.
        The file is created automatically if it does not exist.
      </p>

      <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', marginBottom: '1.5rem', flexWrap: 'wrap' }}>
        {!token && (
          <span style={{ color: '#dc3545', fontSize: '0.9rem' }}>Google account not signed in.</span>
        )}
        {!msAccount && (
          <button onClick={msLoginRedirect} style={btnStyle('#0078d4')}>Sign in with Microsoft</button>
        )}
        {token && msAccount && (
          <button onClick={handleCopy} disabled={running} style={btnStyle('#107c41', running)}>
            {running ? 'Copying…' : 'Copy Google Sheet → OneDrive Excel'}
          </button>
        )}
        {status === 'done' && <span style={{ color: '#107c41', fontWeight: 600 }}>Done!</span>}
        {status === 'error' && <span style={{ color: '#dc3545', fontWeight: 600 }}>Failed — see log below.</span>}
      </div>

      {running && progress && (
        <div style={{ marginBottom: '1rem', padding: '0.5rem 0.75rem', backgroundColor: '#e8f5e9', borderRadius: '4px', fontSize: '0.9rem', color: '#2e7d32' }}>
          {progress}
        </div>
      )}

      {logs.length > 0 && (
        <div style={{ backgroundColor: '#1e1e1e', borderRadius: '6px', padding: '1rem', overflowY: 'auto', maxHeight: '400px' }}>
          {logs.map((line, i) => (
            <div key={i} style={{ fontFamily: 'monospace', fontSize: '0.85rem', color: line.startsWith('✗') ? '#f48771' : line.startsWith('✓') ? '#89d185' : '#d4d4d4', marginBottom: '2px' }}>
              {line}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};
