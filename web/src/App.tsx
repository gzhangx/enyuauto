import { useState } from 'react'
import './App.css'
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { LoginPage } from './components/LoginPage';
import { GoogleSheetsPage } from './components/GoogleSheetsPage';
import { ProjectsPage } from './components/ProjectsPage';

type View = 'projects' | 'sheets';

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
