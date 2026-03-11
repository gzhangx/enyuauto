import { useState } from 'react'
import './App.css'
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { LoginPage } from './components/LoginPage';
import { GoogleSheetsPage } from './components/GoogleSheetsPage';
import { ProjectsPage } from './components/ProjectsPage';
import { MicrosoftOneDrivePage } from './components/MicrosoftOneDrivePage';

type View = 'projects' | 'sheets' | 'onedrive';

const AppContent = () => {
  const { isAuthenticated, logout } = useAuth();
  const [currentView, setCurrentView] = useState<View>('projects');

  if (!isAuthenticated) {
    return <LoginPage />;
  }

  return (
    <>
      <nav style={{
        padding: '1rem',
        backgroundColor: '#f8f9fa',
        borderBottom: '1px solid #dee2e6',
        display: 'flex',
        gap: '1rem',
        alignItems: 'center'
      }}>
        <button
          onClick={() => setCurrentView('projects')}
          style={{
            padding: '0.5rem 1rem',
            backgroundColor: currentView === 'projects' ? '#007bff' : '#6c757d',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer'
          }}
        >
          Projects
        </button>
        <button
          onClick={() => setCurrentView('sheets')}
          style={{
            padding: '0.5rem 1rem',
            backgroundColor: currentView === 'sheets' ? '#007bff' : '#6c757d',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer'
          }}
        >
          Google Sheets
        </button>
        <button
          onClick={() => setCurrentView('onedrive')}
          style={{
            padding: '0.5rem 1rem',
            backgroundColor: currentView === 'onedrive' ? '#0078d4' : '#6c757d',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '0.4rem',
          }}
        >
          <svg width="14" height="14" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <rect x="1" y="1" width="10" height="10" fill="#f25022"/>
            <rect x="12" y="1" width="10" height="10" fill="#7fba00"/>
            <rect x="1" y="12" width="10" height="10" fill="#00a4ef"/>
            <rect x="12" y="12" width="10" height="10" fill="#ffb900"/>
          </svg>
          OneDrive
        </button>
        <div style={{ marginLeft: 'auto' }}>
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
      </nav>
      
      {currentView === 'projects' && <ProjectsPage />}
      {currentView === 'sheets' && <GoogleSheetsPage />}
      {currentView === 'onedrive' && <MicrosoftOneDrivePage />}
    </>
  );
};

function App() {
  return (
    <AuthProvider>
      <AppContent />
    </AuthProvider>
  );
}

export default App
