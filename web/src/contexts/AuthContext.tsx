import type { ISheetInfoSimple } from '@gzhangx/googleapi/lib/googleApi';
import { createContext, useContext, useState, useEffect, type ReactNode } from 'react';
import type { IOpsConfig, ISheetInfoCache } from '../../shared/opsTypes';
import type { ICombinedOpsAndFreeCampData } from '../../shared/main_ops';
import type { AccountInfo } from '@azure/msal-browser';
import { msalInstance, msalReady } from '../lib/msalInstance';
import { readFreedCampCredentials, type FreedCampCredentials } from '../lib/freedCampCredentials';

const GRAPH_SCOPES = ['Files.ReadWrite', 'User.Read'];

export interface IFreedCampCredentials {
  username: string;
  password: string;
}
interface AuthContextType {
  token: string | null;
  msToken: string | null;
  msAccount: AccountInfo | null;
  sheetInfoCache: ISheetInfoCache;
  expiresAt: number | null;
  login: (token: string, expiresIn: number) => void;
  msLogin: (token: string, expiresAt: number) => void;
  msLoginRedirect: () => void;
  msLogout: () => void;
  logout: () => void;
  isAuthenticated: boolean;
  isMsAuthenticated: boolean;
  useMsOps: boolean;
  setUseMsOps: (val: boolean) => void;
  opsConfig: IOpsConfig | null;
  setOpsConfig: (config: IOpsConfig | null) => void;
  combinedOpsAndData: ICombinedOpsAndFreeCampData | null;
  setCombinedOpsAndData: (data: ICombinedOpsAndFreeCampData | null) => void;
  freedCampCredentials: IFreedCampCredentials | null;
  setFreedCampCredentials: (creds: IFreedCampCredentials | null) => void;
  authLoadingStatus: string;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider = ({ children }: { children: ReactNode }) => {
  const [token, setToken] = useState<string | null>(null);
  const [msToken, setMsToken] = useState<string | null>(null);
  const [msAccount, setMsAccount] = useState<AccountInfo | null>(null);
  const [expiresAt, setExpiresAt] = useState<number | null>(null);
  const [sheetInfoCached, setSheetInfoCached] = useState<ISheetInfoSimple[] | null>(null);
  const [opsConfig, setOpsConfig] = useState<IOpsConfig | null>(null);
  const [combinedOpsAndData, setCombinedOpsAndData] = useState<ICombinedOpsAndFreeCampData | null>(null);
  const [freedCampCredentials, setFreedCampCredentials] = useState<{ username: string; password: string } | null>(null);
  const [useMsOps, setUseMsOpsState] = useState<boolean>(() => localStorage.getItem('use_ms_ops') === 'true');
  const [authLoadingStatus, setAuthLoadingStatus] = useState<string>('Initializing Microsoft login...');

  const setUseMsOps = (val: boolean) => {
    setUseMsOpsState(val);
    localStorage.setItem('use_ms_ops', val ? 'true' : 'false');
  };

  // Load token from localStorage on mount, then handle MSAL redirect/silent refresh
  useEffect(() => {
    const savedToken = localStorage.getItem('google_token');
    const savedExpiresAt = localStorage.getItem('google_token_expires_at');
    if (savedToken && savedExpiresAt) {
      const expiresAtNum = parseInt(savedExpiresAt, 10);
      if (Date.now() < expiresAtNum) {
        setToken(savedToken);
        setExpiresAt(expiresAtNum);
      } else {
        localStorage.removeItem('google_token');
        localStorage.removeItem('google_token_expires_at');
      }
    }

    const savedMsToken = localStorage.getItem('ms_token');
    const savedMsExpiresAt = localStorage.getItem('ms_token_expires_at');
    if (savedMsToken && savedMsExpiresAt) {
      const msExpiresAtNum = parseInt(savedMsExpiresAt, 10);
      if (Date.now() < msExpiresAtNum) {
        setMsToken(savedMsToken);
      } else {
        localStorage.removeItem('ms_token');
        localStorage.removeItem('ms_token_expires_at');
      }
    }

    // Handle MSAL redirect result or silent refresh
    msalReady.then(redirectResult => {
      if (redirectResult?.accessToken) {
        const expiresAt = redirectResult.expiresOn?.getTime() ?? Date.now() + 3600 * 1000;
        setMsAccount(redirectResult.account);
        setMsToken(redirectResult.accessToken);
        localStorage.setItem('ms_token', redirectResult.accessToken);
        localStorage.setItem('ms_token_expires_at', expiresAt.toString());
        setAuthLoadingStatus('Setting MS token...');
      } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          setAuthLoadingStatus('Acquiring token silently...');
          msalInstance
            .acquireTokenSilent({ scopes: GRAPH_SCOPES, account: accounts[0] })
            .then(result => {
              if (result?.accessToken) {
                const expiresAt = result.expiresOn?.getTime() ?? Date.now() + 3600 * 1000;
                setMsAccount(accounts[0]);
                setMsToken(result.accessToken);
                localStorage.setItem('ms_token', result.accessToken);
                localStorage.setItem('ms_token_expires_at', expiresAt.toString());
                setAuthLoadingStatus('Setting MS token...');
              } else {
                setAuthLoadingStatus('');
              }
            })
            .catch(() => { setAuthLoadingStatus(''); /* Silent refresh failed – user must log in manually */ });
        } else {
          setAuthLoadingStatus('');
        }
      }
    }).catch(() => { setAuthLoadingStatus(''); });
  }, []);

  // Auto-load FreedCamp credentials once an MS token is available
  useEffect(() => {
    if (!msToken || freedCampCredentials) {
      if (msToken) setAuthLoadingStatus('');
      return;
    }
    setAuthLoadingStatus('Loading FreedCamp credentials...');
    readFreedCampCredentials(msToken)
      .then((creds: FreedCampCredentials | null) => { if (creds) setFreedCampCredentials(creds); })
      .catch(() => {/* non-fatal – user can load manually */})
      .finally(() => { setAuthLoadingStatus(''); });
  }, [msToken]);

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

  const msLogin = (newToken: string, newExpiresAt: number) => {
    setMsToken(newToken);
    localStorage.setItem('ms_token', newToken);
    localStorage.setItem('ms_token_expires_at', newExpiresAt.toString());
  };

  const msLoginRedirect = () => {
    msalReady
      .then(() => msalInstance.loginRedirect({ scopes: GRAPH_SCOPES }))
      .catch(console.error);
  };

  const msLogout = () => {
    setMsToken(null);
    setMsAccount(null);
    localStorage.removeItem('ms_token');
    localStorage.removeItem('ms_token_expires_at');
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalReady
        .then(() => msalInstance.logoutRedirect({ account: accounts[0] }))
        .catch(console.error);
    }
  };

  const logout = () => {
    setToken(null);
    setMsToken(null);
    setMsAccount(null);
    setExpiresAt(null);
    localStorage.removeItem('google_token');
    localStorage.removeItem('google_token_expires_at');
    localStorage.removeItem('ms_token');
    localStorage.removeItem('ms_token_expires_at');

    // Sign out from Microsoft
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalReady
        .then(() => msalInstance.logoutRedirect({ account: accounts[0] }))
        .catch(console.error);
    }
  };

  const isAuthenticated = token !== null && expiresAt !== null && Date.now() < expiresAt;
  const isMsAuthenticated = msToken !== null;

  return (
    <AuthContext.Provider value={{
      token, msToken, msAccount, expiresAt, login, msLogin, msLoginRedirect, msLogout, logout,
      isAuthenticated, isMsAuthenticated, useMsOps, setUseMsOps, authLoadingStatus, sheetInfoCache: {
        getCachedSheetInfo: () => sheetInfoCached,
        setCacheSheetInfo: (data: ISheetInfoSimple[]) => setSheetInfoCached(data),
      },
      opsConfig, setOpsConfig,
      combinedOpsAndData, setCombinedOpsAndData,
      freedCampCredentials, setFreedCampCredentials,
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
