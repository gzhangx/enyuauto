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
  const [folderInput, setFolderInput] = useState('');
  const [files, setFiles] = useState<DriveItem[] | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

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
        files.length === 0 ? (
          <p style={{ color: '#777' }}>No files found in this folder.</p>
        ) : (
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr>
                {['Name', 'Type', 'Size', 'Last Modified'].map(h => (
                  <th key={h} style={{
                    border: '1px solid #ddd',
                    padding: '10px 12px',
                    textAlign: 'left',
                    backgroundColor: '#f8f9fa',
                    position: 'sticky',
                    top: 0,
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {files.map(item => (
                <tr key={item.id}>
                  <td style={{ border: '1px solid #ddd', padding: '10px 12px' }}>
                    {item.webUrl
                      ? <a href={item.webUrl} target="_blank" rel="noreferrer">{item.name}</a>
                      : item.name}
                  </td>
                  <td style={{ border: '1px solid #ddd', padding: '10px 12px' }}>
                    {item.folder ? '\U0001f4c1 Folder' : (item.file?.mimeType ?? 'File')}
                  </td>
                  <td style={{ border: '1px solid #ddd', padding: '10px 12px' }}>
                    {item.size != null
                      ? item.size < 1024
                        ? `${item.size} B`
                        : item.size < 1024 * 1024
                          ? `${(item.size / 1024).toFixed(1)} KB`
                          : `${(item.size / (1024 * 1024)).toFixed(1)} MB`
                      : '\u2014'}
                  </td>
                  <td style={{ border: '1px solid #ddd', padding: '10px 12px' }}>
                    {item.lastModifiedDateTime
                      ? new Date(item.lastModifiedDateTime).toLocaleString()
                      : '\u2014'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )
      )}
    </div>
  );
};
