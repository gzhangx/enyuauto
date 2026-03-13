import { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';

export const GoogleSheetsPage = () => {
  const { token, logout } = useAuth();
  const [sheetId, setSheetId] = useState('');
  const [sheetData, setSheetData] = useState<any[][]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');

  const readSheet = async () => {
    if (!sheetId.trim()) {
      setError('Please enter a Google Sheet ID');
      return;
    }

    setLoading(true);
    setError('');
    setSheetData([]);

    try {
      // Using Google Sheets API v4
      const range = 'Sheet1'; // You can modify this or make it configurable
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${range}`;
      
      const response = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || 'Failed to fetch sheet data');
      }

      const data = await response.json();
      setSheetData(data.values || []);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An error occurred');
      console.error('Error reading sheet:', err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '2rem', maxWidth: '1200px', margin: '0 auto' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '2rem' }}>
        <h1>Google Sheets Reader</h1>
        <button 
          onClick={logout}
          style={{
            padding: '0.5rem 1rem',
            backgroundColor: '#dc3545',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer'
          }}
        >
          Logout
        </button>
      </div>

      <div style={{ marginBottom: '2rem' }}>
        <div style={{ display: 'flex', gap: '1rem', marginBottom: '1rem' }}>
          <input
            type="text"
            placeholder="Enter Google Sheet ID"
            value={sheetId}
            onChange={(e) => setSheetId(e.target.value)}
            style={{
              flex: 1,
              padding: '0.5rem',
              border: '1px solid #ddd',
              borderRadius: '4px'
            }}
          />
          <button
            onClick={readSheet}
            disabled={loading}
            style={{
              padding: '0.5rem 1.5rem',
              backgroundColor: loading ? '#6c757d' : '#007bff',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: loading ? 'not-allowed' : 'pointer'
            }}
          >
            {loading ? 'Loading...' : 'Read Sheet'}
          </button>
        </div>
        <p style={{ fontSize: '0.875rem', color: '#6c757d' }}>
          Enter the Sheet ID from the URL: https://docs.google.com/spreadsheets/d/<strong>SHEET_ID</strong>/edit
        </p>
      </div>

      {error && (
        <div style={{
          padding: '1rem',
          backgroundColor: '#f8d7da',
          color: '#721c24',
          border: '1px solid #f5c6cb',
          borderRadius: '4px',
          marginBottom: '1rem'
        }}>
          {error}
        </div>
      )}

      {sheetData.length > 0 && (
        <div style={{ overflowX: 'auto' }}>
          <table style={{
            width: '100%',
            borderCollapse: 'collapse',
            border: '1px solid #ddd'
          }}>
            <tbody>
              {sheetData.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, cellIndex) => (
                    <td
                      key={cellIndex}
                      style={{
                        border: '1px solid #ddd',
                        padding: '8px',
                        backgroundColor: rowIndex === 0 ? '#f2f2f2' : 'white'
                      }}
                    >
                      {cell}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};
