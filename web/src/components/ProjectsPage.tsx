import { useEffect, useState, useRef, useCallback } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import {
  getFreeCampAndUpdateOperations, getOpsAndMainList, getSheetOps, processOperation,deleteItemActionTask,
  type FreeCampAndUpdateOperations,
  loadMainData,
  combineOpsConfigWithFreedCampData,
  type DebugLog,
  completeDateUpdater,  
  updateDoneParentIds,
  type IOneDriveDirInfo,
  createMsExcelDataOps,
} from '../../shared/main_ops';
import { ErrorDialog } from './ErrorDialog';
import { freedCampOps } from '../lib/util';
import type { FreedCampLoginParams, LoginResponse } from '../../shared/freedcampTypes';
import { type IOperationWithLineNumberAndParentTaskId } from '../../shared/opsTypes';
import { renderActionCell, renderSyncActionCell, type RenderActionCellDeps } from './projectsPageUtil/render_action_cell';
import { fetchOneDriveChildren } from './MicrosoftOneDrivePage';


type LogMessage = {
  id: number;
  text: string;
};

type UseLoggerResult = {
  logMessages: LogMessage[];
  criticalError: { show: boolean; message: string };
  closeCriticalError: () => void;
} & DebugLog;

const useLogger = (displayDuration = 60000): UseLoggerResult => {
  const [logMessages, setLogMessages] = useState<LogMessage[]>([]);
  const [criticalError, setCriticalError] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const logIdCounterRef = useRef(0);
  
  const doLog = useCallback<DebugLog['doLog']>((msg, critical = false) => {
    if (critical) {
      console.error(msg);
      setCriticalError({ show: true, message: msg });
    } else {
      console.log(msg);
    }
    const newLog: LogMessage = {
      id: logIdCounterRef.current++,
      text: msg
    };
    setLogMessages(prev => {
      const next = [...prev, newLog];
      return next.length > 50 ? next.slice(next.length - 50) : next;
    });

    setTimeout(() => {
      setLogMessages(prev => prev.filter(log => log.id !== newLog.id));
    }, displayDuration);
  }, [displayDuration]);

  const closeCriticalError = useCallback(() => {
    setCriticalError({ show: false, message: '' });
  }, []);

  return { logMessages, doLog, criticalError, closeCriticalError };
};


export const ProjectsPage = ({ onNavigateToFreedCamp }: { onNavigateToFreedCamp?: () => void }) => {
  const { token, msToken, msAccount, msLoginRedirect, msLogout, sheetInfoCache, opsConfig, setOpsConfig, combinedOpsAndData, setCombinedOpsAndData, freedCampCredentials, useMsOps } = useAuth();  
  const [fullProjectList, setFullProjectList] = useState<IOperationWithLineNumberAndParentTaskId[]>([]);
  const [projectList, setProjectList] = useState<IOperationWithLineNumberAndParentTaskId[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  const [isLoading, setIsLoading] = useState({
    listLoading: '',
    freeCampLoading: '',
    projectButtonAction: '',
  });
  const [progressText, setProgressText] = useState('');
  const [errorDialog, setErrorDialog] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const sheetOpsRef = useRef<Awaited<ReturnType<typeof getSheetOps>> | null>(null);
  const { logMessages, doLog, criticalError, closeCriticalError } = useLogger(60000);
  const [showLogPanel, setShowLogPanel] = useState(false);
  const [cachedToken, setCachedToken] = useState<{ token: LoginResponse; timestamp: number } | null>(null);
  const [showAllProjects, setShowAllProjects] = useState(false);
  const [daysShowAlertAfterComplete, setDaysShowAlertAfterComplete] = useState(3);
  
  const originalFreedCampOps = getFreeCampAndUpdateOperations(freedCampOps);
  

  const startupRunOnce = useRef(false);
  const freeCampOpsWithCache: FreeCampAndUpdateOperations = {
    ...originalFreedCampOps,
    getFreedCampToken: async (opsAndTemplates) => {
      const now = Date.now();
      const oneHour = 60 * 60 * 1000; // 1 hour in milliseconds
      
      // Check if cached token exists and is still valid
      if (cachedToken && (now - cachedToken.timestamp) < oneHour) {
        doLog('Using cached FreedCamp token');
        return cachedToken.token;
      }
      
      // Get new token if cache is expired or doesn't exist
      doLog('Fetching new FreedCamp token');
      const newToken = await originalFreedCampOps.getFreedCampToken(opsAndTemplates);
      setCachedToken({ token: newToken, timestamp: now });
      return newToken;
    },    
  };

  const logParam = { getOperations: () => [], doLog }
  async function fetchData(freedCampCredentials: FreedCampLoginParams, fullReload = false) {
    console.log('useEffect run');
    const activeToken = useMsOps ? msToken : token;
    if (activeToken) {
      const ops = useMsOps
        ? createMsExcelDataOps(msToken!)
        : await getSheetOps({ token: token! }, sheetInfoCache);
      if (opsConfig &&combinedOpsAndData && !fullReload) {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects only...' }));
        const res = await loadMainData(ops);
        updateDoneParentIds(combinedOpsAndData, res.operationList, logParam);
        prepareData(res);
        setIsLoading(prev=>({ ...prev, listLoading: '' }));
      } else {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects and config...' }));
      
        
        await getOpsAndMainList(ops, logParam).then(async res => {
          setOpsConfig(res); // Save the complete response (including functions)
          //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
          const combined = await combineOpsConfigWithFreedCampData(res, freeCampOpsWithCache, freedCampCredentials, logParam);          

          //--------------------------- get freecamp data        
          const pr = freeCampOpsWithCache.getFreedCampProcessor(combined.loginToken);
          for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
            const actionInfo = combined.opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
            if (actionInfo) {
              setIsLoading(prev => ({ ...prev, freeCampLoading: `Loading FreeCamp tasks for ${action}...` }));
              const prjs = await pr.getTasksForProjects(actionInfo.project_id, 1);
              combined.freedCampTasksByAction[action] = prjs;
            }
          }
          setCombinedOpsAndData(combined);
          updateDoneParentIds(combined, res.operationList, logParam);
          setIsLoading(prev => ({ ...prev, freeCampLoading: '' }));      
          //---------------------------
          prepareData(res);
        }).catch(error => {
          console.error('Error loading projects:', error);
          setResponseData(`Error: ${error.message || String(error)}`);
          doLog(`Error: ${error.message || String(error)}`);
          setErrorDialog({ show: true, message: `Failed to load projects:\n${error.message || String(error)}` });
        }).finally(() => {
          setIsLoading(prev => ({ ...prev, listLoading: '' }));
        });
      }
      sheetOpsRef.current = ops;
    }

    function prepareData(dataRes: { operationList: IOperationWithLineNumberAndParentTaskId[];}) {
      const list = dataRes.operationList.map((item, index) => ({ ...item, line: index + 2 }));
      setFullProjectList(list);
    }
  }

  useEffect(() => {
    const filtered = fullProjectList.filter((item) => {
      if (showAllProjects) {
        return true;
      }
      return !item.isFinished;
    });
    setProjectList(filtered);
  }, [fullProjectList, showAllProjects]);

  //useEffect(() => {
    //if (!freedCampCredentials && onNavigateToFreedCamp) {
      //onNavigateToFreedCamp();
    //}
  //}, [freedCampCredentials]);

  useEffect(() => {
    const ready = useMsOps ? !!msToken : !!token;
    if (ready && !startupRunOnce.current && freedCampCredentials) {
      startupRunOnce.current = true;
      fetchData(freedCampCredentials);
    }
  }, [token, msToken, freedCampCredentials, useMsOps]);
  


  // ...function removed, now imported from render_action_cell.tsx

  // Calculate animation speed based on number of messages
  if (isLoading.listLoading || isLoading.freeCampLoading || isLoading.projectButtonAction) {
    return (
      <div style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '100vh',
        gap: '20px'
      }}>
        <div style={{
          width: '60px',
          height: '60px',
          border: '6px solid rgba(33, 150, 243, 0.2)',
          borderTop: '6px solid rgb(33, 150, 243)',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite'
        }}></div>
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.listLoading}</h2>
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.freeCampLoading}</h2>
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading.projectButtonAction}</h2>
        {progressText && (
          <div style={{ 
            maxWidth: '600px', 
            padding: '10px 20px',
            color: '#777',
            fontSize: '13px',
            textAlign: 'center',
            lineHeight: '1.4'
          }}>
            {progressText}
          </div>
        )}
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}</style>

      </div>
    );
  }

  const actionCellDeps: RenderActionCellDeps = {
    sheetOpsRef,
    combined: combinedOpsAndData,
    freeCampOpsWithCache,
    freedCampLoginParams: freedCampCredentials!,
    deleteItemActionTask,
    completeDateUpdater,
    doLog,
    fetchData,
    setErrorDialog,
    daysShowAlertAfterComplete,
  };

  const stickyHeaderCellStyle = {
    border: '1px solid #ddd',
    padding: '12px',
    textAlign: 'left' as const,
    position: 'sticky' as const,
    top: 0,
    zIndex: 900,
    backgroundColor: '#fff'
  };

  if (!freedCampCredentials) {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: '40vh', gap: '1rem' }}>
        <p style={{ color: '#555' }}>FreedCamp credentials not loaded. Please set them first.</p>
        {onNavigateToFreedCamp && (
          <button
            onClick={onNavigateToFreedCamp}
            style={{ padding: '0.5rem 1.25rem', backgroundColor: '#e65100', color: 'white', border: 'none', borderRadius: '4px', fontSize: '0.95rem', cursor: 'pointer' }}
          >
            Go to FreedCamp Login
          </button>
        )}
      </div>
    );
  }
  return (
    <>
      <ErrorDialog 
        show={errorDialog.show}
        message={errorDialog.message}
        onClose={() => setErrorDialog({ show: false, message: '' })}
      />

      <ErrorDialog 
        show={criticalError.show}
        message={criticalError.message}
        onClose={closeCriticalError}
      />

      <button
        type="button"
        aria-label={showLogPanel ? 'Hide logs' : 'Show logs'}
        onClick={() => setShowLogPanel(prev => !prev)}
        style={{
          position: 'fixed',
          top: '12px',
          right: '12px',
          width: '36px',
          height: '36px',
          borderRadius: '18px',
          border: '1px solid #ddd',
          backgroundColor: '#ffffff',
          color: '#333',
          fontSize: '12px',
          fontWeight: 700,
          cursor: 'pointer',
          zIndex: 1200,
          boxShadow: '0 2px 6px rgba(0,0,0,0.15)'
        }}
      >
        {logMessages.length}
      </button>

      {showLogPanel && (
      <div style={{
        position: 'fixed',
        top: '56px',
        right: '12px',
        width: '340px',
        height: '70vh',
        backgroundColor: '#fff',
        border: '1px solid #ddd',
        borderRadius: '6px',
        overflowY: 'auto',
        zIndex: 1100,
        display: 'flex',
        flexDirection: 'column',
        gap: '6px',
        padding: '10px',
        boxShadow: '0 4px 12px rgba(0,0,0,0.2)'
      }}>
        {logMessages.length === 0 ? (
          <div style={{ color: '#777', fontSize: '12px' }}>No logs yet</div>
        ) : (
          [...logMessages].reverse().map((log) => (
            <div
              key={log.id}
              style={{
                backgroundColor: 'rgba(33, 150, 243, 0.9)',
                color: 'white',
                padding: '8px 12px',
                borderRadius: '4px',
                fontSize: '12px',
                wordWrap: 'break-word',
                boxShadow: '0 2px 4px rgba(0,0,0,0.2)'
              }}
            >
              {log.text}
            </div>
          ))
        )}
      </div>
      )}

      <style>{`
        @keyframes slideInFromBottom {
          from {
            transform: translateY(100%);
            opacity: 0;
          }
          to {
            transform: translateY(0);
            opacity: 1;
          }
        }
      `}</style>

      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px' }}>
        <h1 style={{ margin: 0 }}>Enyu Site</h1>
        <button className="btn btn-create" onClick={() => fetchData(freedCampCredentials,false)}>Reload</button>
        <button className="btn btn-create" onClick={()=>fetchData(freedCampCredentials, true)}>Reload All</button>
        {msAccount && (
          <button className="btn btn-create" onClick={msLogout} style={{ backgroundColor: '#e3f2fd', color: '#1565c0' }}>
            MS Logout ({msAccount.username})
          </button>
        )}
        {!msAccount && (
          <button className="btn btn-create" onClick={msLoginRedirect} style={{ backgroundColor: '#e3f2fd', color: '#1565c0' }}>
            MS Login
          </button>
        )}
        <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '14px' }}>
          <input
            type="checkbox"
            checked={showAllProjects}
            onChange={(e) => setShowAllProjects(e.target.checked)}
          />
          Show all
        </label>
        <input
          type="number"
          min={0}
          max={10}
          placeholder="days"
          value={daysShowAlertAfterComplete}
          onChange={(e) => {
            const v = parseInt(e.target.value, 10);
            if (!isNaN(v) && v >= 0 && v <= 10) setDaysShowAlertAfterComplete(v);
          }}
          style={{ width: '60px', fontSize: '14px', padding: '2px 6px' }}
        />
      </div>
      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '2rem' }}>
        <thead>
          <tr>
            <th style={stickyHeaderCellStyle}>Action</th>
            <th style={stickyHeaderCellStyle}>文件</th>
            <th style={stickyHeaderCellStyle}>文章名</th>
            <th style={stickyHeaderCellStyle}>作者</th>
            {opsConfig?.groupAndMainProjectMapping.actions.map((action) => (
              <th key={action} style={stickyHeaderCellStyle}>
                {action}
              </th>
            ))}
            <th style={stickyHeaderCellStyle}>Line</th>
          </tr>
        </thead>
        <tbody>
          {
            projectList.map((p,idx) => {
              return <tr key={p.文章名+idx}>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.noNeedToCreate === false && <button className="btn btn-create" onClick={
                  async () => {
                    if (opsConfig === null) {
                      setResponseData('Operation data not loaded yet.');
                      return;
                    }
                    setIsLoading(prev => ({ ...prev, projectButtonAction: 'Creating project...' }));
                    try {
                      const ops = sheetOpsRef.current;
                      if (ops) {
                        const logs: DebugLog = {
                          doLog: msg => {
                            console.log(msg);
                            setProgressText(msg);
                          }
                        };
                        if (combinedOpsAndData) {
                          let oneDriveFolders: IOneDriveDirInfo[] = [];
                          if (p.mainFolder) {
                            if (!msToken) {
                              setErrorDialog({ show: true, message: 'Microsoft token not available. Please sign in to Microsoft to access OneDrive folders.' });
                              return;
                            }
                            console.log('debugremove sending main folder', p.mainFolder)
                            oneDriveFolders = await fetchOneDriveChildren(msToken, p.mainFolder);
                          }
                          await processOperation(ops, freeCampOpsWithCache, combinedOpsAndData, p, oneDriveFolders, logs);
                        }
                      }
                    } catch (error: any) {
                      console.error('Error creating project:', error);                                            
                      setErrorDialog({ show: true, message: `Failed to create project:\n${error.message || String(error)}` });
                      return;
                    } finally {
                      setIsLoading(prev => ({ ...prev, projectButtonAction: '' }));
                      setProgressText('');
                    }
                    //const retData = await createOrDelProject(p.itemPositionOnSheet, 'main');
                    //setResponseData(JSON.stringify(retData, null, 2));
                  }
                }>Create</button>}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>
                  <a href={p.文件} target="_blank">{p.文件}</a>
                </td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.文章名}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.作者}</td>
                {opsConfig?.groupAndMainProjectMapping.actions.map((action) => (
                  <td key={action} style={{ border: '1px solid #ddd', padding: '12px' }}>
                    {renderActionCell(p, action, actionCellDeps)}
                  </td>
                ))}
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.itemPositionOnSheet}
                  { renderSyncActionCell(p, freedCampCredentials, actionCellDeps, 'Sync from FreedCamp') }
                </td>
              </tr>
            })
          }
        </tbody>
      </table>      
      {responseData && (
        <div style={{ marginTop: '2rem', textAlign: 'left' }}>
          <h2>Response Data:</h2>
          <pre style={{ 
            backgroundColor: '#f5f5f5', 
            padding: '1rem', 
            borderRadius: '4px',
            overflow: 'auto',
            maxHeight: '400px'
          }}>
            {responseData}
          </pre>
        </div>
      )}
      <p className="read-the-docs">
        
      </p>
    </>
  )
}
