import { useGoogleLogin } from '@react-oauth/google';
import { useAuth } from '../contexts/AuthContext';

export const LoginPage = ({ onGoToOneDrive }: { onGoToOneDrive?: () => void }) => {
  const { login, msLoginRedirect, isMsAuthenticated, msLogout, msAccount, useMsOps, setUseMsOps } = useAuth();

  const googleLogin = useGoogleLogin({
    onSuccess: (tokenResponse) => {
      // tokenResponse contains access_token and expires_in
      login(tokenResponse.access_token, tokenResponse.expires_in);
    },
    onError: () => {
      console.error('Login Failed');
      alert('Google login failed. Please try again.');
    },
    scope: 'https://www.googleapis.com/auth/spreadsheets',
  });

  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      minHeight: '100vh',
      gap: '2rem'
    }}>
      <h1>Welcome to Enyu Auto</h1>
      <p>Please sign in to continue</p>
      <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '1rem', cursor: 'pointer' }}>
        <input
          type="checkbox"
          checked={useMsOps}
          onChange={e => setUseMsOps(e.target.checked)}
          style={{ width: '16px', height: '16px', cursor: 'pointer' }}
        />
        Use Microsoft (OneDrive Excel)
      </label>
      <button
        onClick={() => googleLogin()}
        style={{
          padding: '0.75rem 2rem',
          backgroundColor: '#4285f4',
          color: 'white',
          border: 'none',
          borderRadius: '4px',
          fontSize: '1rem',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          gap: '0.5rem'
        }}
      >
        <svg width="18" height="18" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
          <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
          <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
          <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
          <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
          <path fill="none" d="M0 0h48v48H0z"/>
        </svg>
        Sign in with Google
      </button>
      <button
        onClick={msLoginRedirect}
        style={{
          padding: '0.75rem 2rem',
          backgroundColor: '#0078d4',
          color: 'white',
          border: 'none',
          borderRadius: '4px',
          fontSize: '1rem',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          gap: '0.5rem'
        }}
      >
        <svg width="18" height="18" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
          <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
          <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
          <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
          <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
        </svg>
        Sign in with Microsoft
      </button>
      {isMsAuthenticated && (
        <>
          {msAccount && (
            <div style={{ textAlign: 'center', lineHeight: '1.4' }}>
              <div style={{ fontWeight: 600, fontSize: '1rem', color: '#212529' }}>{msAccount.name}</div>
              <div style={{ fontSize: '0.85rem', color: '#6c757d' }}>{msAccount.username}</div>
            </div>
          )}
          <button
            onClick={onGoToOneDrive}
            style={{
              padding: '0.75rem 2rem',
              backgroundColor: '#0078d4',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              fontSize: '1rem',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '0.5rem'
            }}
          >
            <svg width="18" height="18" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
              <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
              <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
              <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
              <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
            </svg>
            Go to Microsoft OneDrive
          </button>
          <button
            onClick={msLogout}
            style={{
              padding: '0.75rem 2rem',
              backgroundColor: '#6c757d',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              fontSize: '1rem',
              cursor: 'pointer'
            }}
          >
            Logout Microsoft
          </button>
        </>
      )}
    </div>
  );
};
