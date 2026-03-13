import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';

type DriveItem = {
  id: string;
  name: string;
  size?: number;
  webUrl: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  lastModifiedDateTime?: string;
  parentReference?: {
    driveId?: string;
    id?: string;
    path?: string;
    name: string;
    siteId: string;
    driveType: string; //'documentLibrary' | 'personal' | 'business' | 'unknownFutureValue';
  }
};

function encodeSharingUrl(url: string): string {
  const b64 = btoa(url).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  return 'u!' + b64;
}

export async function fetchOneDriveChildren(token: string, input: string): Promise<DriveItem[]> {
  const trimmed = input.trim().replace(/\/?$/, '');
  let url: string;
  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
    const encoded = encodeSharingUrl(trimmed);
    url = `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem/children`;
  } else if (trimmed === '') {
    url = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
  } else {
    const path = trimmed.replace(/^\/+/, '');
    url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path)}:/children`;
  }
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body?.error?.message || res.statusText);
  }
  const data = await res.json();
  return data.value as DriveItem[];
}

async function resolveToWorkbookBaseUrl(token: string, input: string): Promise<string> {
  const trimmed = input.trim();
  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
    // Resolve sharing URL → get driveId + itemId, then use /drives/{driveId}/items/{itemId}
    const encoded = encodeSharingUrl(trimmed);
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem?$select=id,parentReference`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(body?.error?.message || res.statusText);
    }
    const item = await res.json();
    const driveId = item.parentReference?.driveId;
    const itemId = item.id;
    if (!driveId || !itemId) throw new Error('Could not resolve driveId/itemId from sharing link');
    return `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`;
  } else {
    return `https://graph.microsoft.com/v1.0/me/drive/items/${trimmed}`;
  }
}

async function readXlsxSheet(token: string, input: string, sheetName = 'Sheet1'): Promise<string[][]> {
  const baseUrl = await resolveToWorkbookBaseUrl(token, input);
  const headers: Record<string, string> = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };

  // Create a read-only workbook session (required for AAD/SharePoint files)
  let sessionId: string | null = null;
  try {
    const sessionRes = await fetch(`${baseUrl}/workbook/createSession`, {
      method: 'POST',
      headers,
      body: JSON.stringify({ persistChanges: false }),
    });
    if (sessionRes.ok) {
      const sessionData = await sessionRes.json();
      sessionId = sessionData.id ?? null;
    }
  } catch {
    // If session creation fails, try without session
  }

  const readHeaders: Record<string, string> = { Authorization: `Bearer ${token}` };
  if (sessionId) readHeaders['workbook-session-id'] = sessionId;

  // Step 0: fetch only row 1 (up to column AZ) to detect actual column count
  const usedRangeRes = await fetch(
    `${baseUrl}/workbook/worksheets/${encodeURIComponent(sheetName)}/usedRange`,
    { headers: readHeaders }
  );
  if (!usedRangeRes.ok) {
    const body = await usedRangeRes.json().catch(() => ({}));
    throw new Error(body?.error?.message || usedRangeRes.statusText);
  }

  const data = await usedRangeRes.json();

  // Close session (fire and forget)
  if (sessionId) {
    fetch(`${baseUrl}/workbook/closeSession`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'workbook-session-id': sessionId },
    }).catch(() => {});
  }

  const values: string[][] = (data.values ?? []) as string[][];
  // Trim trailing empty rows
  const lastNonEmpty = values.reduceRight((acc, row, i) => acc !== -1 ? acc : row.some(c => c !== null && c !== '') ? i : -1, -1);
  return lastNonEmpty === -1 ? [] : values.slice(0, lastNonEmpty + 1);
}

const btnStyle = (color: string): React.CSSProperties => ({
  padding: '0.5rem 1.25rem',
  backgroundColor: color,
  color: 'white',
  border: 'none',
  borderRadius: '4px',
  fontSize: '0.95rem',
  cursor: 'pointer',
});

export const MicrosoftOneDrivePage = () => {
  const { msToken, msAccount, msLoginRedirect } = useAuth();
  const [folderInput, setFolderInput] = useState('https://acccnusa.sharepoint.com/sites/enyueditors/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2Fenyueditors%2FShared%20Documents%2FGeneral%2FEnYu%202026%2F2026%2D03%2F6%5F%E5%AE%A3%E6%95%99%E6%97%A0%E5%9B%BD%E7%95%8C%202026%2D3%2D1%20Pauline%20Sattles%20Feb%202026%20newsletter%20Joy&viewid=8059c88a%2D1194%2D4065%2Db31a%2D3749f1c293f2');
  const [files, setFiles] = useState<DriveItem[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [xlsxInput, setXlsxInput] = useState('');
  const [xlsxSheetName, setXlsxSheetName] = useState('Sheet1');
  const [xlsxData, setXlsxData] = useState<string[][] | null>(null);
  const [xlsxLoading, setXlsxLoading] = useState(false);
  const [xlsxError, setXlsxError] = useState('');

  // On mount: consume redirect result OR silently refresh a cached session
  // (This is now handled centrally by AuthContext)

  const handleLogin = () => {
    msLoginRedirect();
  };

  const handleListFiles = async () => {
    if (!msToken || !msAccount) return;
    setLoading(true);
    setError('');
    setFiles(null);
    try {
      console.log('debugremove sending main folder', folderInput)
      const items = await fetchOneDriveChildren(msToken, folderInput);
      setFiles(items);
    } catch (err: any) {
      setError(err.message || String(err));
    } finally {
      setLoading(false);
    }
  };

  const handleReadXlsx = async () => {
    if (!msToken || !xlsxInput.trim()) return;
    setXlsxLoading(true);
    setXlsxError('');
    setXlsxData(null);
    try {
      const rows = await readXlsxSheet(msToken, xlsxInput, xlsxSheetName || 'Sheet1');
      setXlsxData(rows);
    } catch (err: any) {
      setXlsxError(err.message || String(err));
    } finally {
      setXlsxLoading(false);
    }
  };

  if (!msAccount || !msToken) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '60vh',
        gap: '1.5rem',
      }}>
        <h2 style={{ margin: 0 }}>Microsoft OneDrive</h2>
        <p style={{ color: '#555', margin: 0 }}>
          Sign in with your Microsoft account to browse OneDrive files.
        </p>
        <button onClick={handleLogin} style={{
          ...btnStyle('#0078d4'),
          display: 'flex',
          alignItems: 'center',
          gap: '0.5rem',
          fontSize: '1rem',
          padding: '0.75rem 2rem',
        }}>
          <svg width="18" height="18" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
            <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
            <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
            <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
          </svg>
          Sign in with Microsoft
        </button>
        {error && <p style={{ color: '#dc3545', margin: 0 }}>{error}</p>}
      </div>
    );
  }

  return (
    <div style={{ padding: '1.5rem', maxWidth: '900px', margin: '0 auto' }}>
      <h2 style={{ marginBottom: '1rem' }}>Microsoft OneDrive Browser</h2>

      <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '1rem' }}>
        <input
          type="text"
          value={folderInput}
          onChange={e => setFolderInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && handleListFiles()}
          placeholder="Paste a OneDrive folder link, or type a path (e.g. Documents/Reports)"
          style={{
            flex: 1,
            padding: '0.5rem 0.75rem',
            fontSize: '0.95rem',
            border: '1px solid #ccc',
            borderRadius: '4px',
          }}
        />
        <button onClick={handleListFiles} disabled={loading} style={btnStyle('#0078d4')}>
          {loading ? 'Loading\u2026' : 'List Files'}
        </button>
      </div>

      {error && (
        <p style={{ color: '#dc3545', marginBottom: '1rem' }}>{error}</p>
      )}

      {files !== null && (
        // folder listing table below
        <></>
      )}

      <hr style={{ margin: '1.5rem 0', borderColor: '#ddd' }} />
      <h3 style={{ marginBottom: '0.75rem' }}>Read Excel Sheet</h3>
      <div style={{ display: 'flex', gap: '0.5rem', marginBottom: '0.5rem', flexWrap: 'wrap' }}>
        <input
          type="text"
          value={xlsxInput}
          onChange={e => setXlsxInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && handleReadXlsx()}
          placeholder="Paste a OneDrive xlsx link or item ID"
          style={{ flex: 3, minWidth: '220px', padding: '0.5rem 0.75rem', fontSize: '0.95rem', border: '1px solid #ccc', borderRadius: '4px' }}
        />
        <input
          type="text"
          value={xlsxSheetName}
          onChange={e => setXlsxSheetName(e.target.value)}
          placeholder="Sheet name (default: Sheet1)"
          style={{ flex: 1, minWidth: '140px', padding: '0.5rem 0.75rem', fontSize: '0.95rem', border: '1px solid #ccc', borderRadius: '4px' }}
        />
        <button onClick={handleReadXlsx} disabled={xlsxLoading} style={btnStyle('#107c41')}>
          {xlsxLoading ? 'Loading…' : 'Read Sheet'}
        </button>
      </div>
      {xlsxError && <p style={{ color: '#dc3545', marginBottom: '1rem' }}>{xlsxError}</p>}
      {xlsxData !== null && (
        xlsxData.length === 0 ? (
          <p style={{ color: '#777' }}>Sheet is empty.</p>
        ) : (
          <div style={{ overflowX: 'auto', marginTop: '0.75rem' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
              <thead>
                <tr>
                  {xlsxData[0].map((cell, ci) => (
                    <th key={ci} style={{ border: '1px solid #ddd', padding: '8px 10px', backgroundColor: '#f8f9fa', textAlign: 'left', position: 'sticky', top: 0 }}>
                      {cell}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {xlsxData.slice(1).map((row, ri) => (
                  <tr key={ri}>
                    {row.map((cell, ci) => (
                      <td key={ci} style={{ border: '1px solid #ddd', padding: '8px 10px' }}>{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )
      )}

    </div>
  );
};
