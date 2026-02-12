import type { ISheetInfoSimple } from '@gzhangx/googleapi/lib/googleApi';
import { createContext, useContext, useState, useEffect, type ReactNode } from 'react';
import type { IOpsConfig, ISheetInfoCache } from '../../shared/opsTypes';
import type { ICombinedOpsAndFreeCampData } from '../../shared/main_ops';



interface AuthContextType {
  token: string | null;
  sheetInfoCache: ISheetInfoCache;
  expiresAt: number | null;
  login: (token: string, expiresIn: number) => void;
  logout: () => void;
  isAuthenticated: boolean;
  opsConfig: IOpsConfig | null;
  setOpsConfig: (config: IOpsConfig | null) => void;
  combinedOpsAndData: ICombinedOpsAndFreeCampData | null;
  setCombinedOpsAndData: (data: ICombinedOpsAndFreeCampData | null) => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider = ({ children }: { children: ReactNode }) => {
  const [token, setToken] = useState<string | null>(null);
  const [expiresAt, setExpiresAt] = useState<number | null>(null);
  const [sheetInfoCached, setSheetInfoCached] = useState<ISheetInfoSimple[] | null>(null);
  const [opsConfig, setOpsConfig] = useState<IOpsConfig | null>(null);
  const [combinedOpsAndData, setCombinedOpsAndData] = useState<ICombinedOpsAndFreeCampData | null>(null);
  // Load token from localStorage on mount
  useEffect(() => {
    const savedToken = localStorage.getItem('google_token');
    const savedExpiresAt = localStorage.getItem('google_token_expires_at');
    
    if (savedToken && savedExpiresAt) {
      const expiresAtNum = parseInt(savedExpiresAt, 10);
      // Check if token is still valid
      if (Date.now() < expiresAtNum) {
        setToken(savedToken);
        setExpiresAt(expiresAtNum);
      } else {
        // Token expired, clear it
        localStorage.removeItem('google_token');
        localStorage.removeItem('google_token_expires_at');
      }
    }
  }, []);

  // Auto-logout when token expires
  useEffect(() => {
    if (expiresAt) {
      const timeUntilExpiry = expiresAt - Date.now();
      if (timeUntilExpiry > 0) {
        const timer = setTimeout(() => {
          logout();
        }, timeUntilExpiry);
        return () => clearTimeout(timer);
      }
    }
  }, [expiresAt]);

  const login = (newToken: string, expiresIn: number) => {
    const expiresAtTime = Date.now() + expiresIn * 1000;
    setToken(newToken);
    setExpiresAt(expiresAtTime);
    localStorage.setItem('google_token', newToken);
    localStorage.setItem('google_token_expires_at', expiresAtTime.toString());
  };

  const logout = () => {
    setToken(null);
    setExpiresAt(null);
    localStorage.removeItem('google_token');
    localStorage.removeItem('google_token_expires_at');
  };

  const isAuthenticated = token !== null && expiresAt !== null && Date.now() < expiresAt;

  return (
    <AuthContext.Provider value={{
      token, expiresAt, login, logout, isAuthenticated, sheetInfoCache: {
        getCachedSheetInfo: () => sheetInfoCached,
        setCacheSheetInfo: (data: ISheetInfoSimple[]) => setSheetInfoCached(data),
      },
      opsConfig, setOpsConfig,
      combinedOpsAndData, setCombinedOpsAndData,
    }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within an AuthProvider');
  }
  return context;
};
