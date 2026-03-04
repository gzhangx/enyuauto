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
} from '../../shared/main_ops';
import { ErrorDialog } from './ErrorDialog';
import { freedCampOps } from '../lib/util';
import type { LoginResponse } from '../../shared/freedcampTypes';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig } from '../../shared/opsTypes';
import { renderActionCell, renderSyncActionCell, type RenderActionCellDeps } from './projectsPageUtil/render_action_cell';


type LogMessage = {
  id: number;
  text: string;
};

const useLogger = (displayDuration = 60000) => {
  const [logMessages, setLogMessages] = useState<LogMessage[]>([]);
  const logIdCounterRef = useRef(0);
  
  const doLog = useCallback((msg: string) => {
    console.log(msg);
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

  return { logMessages, doLog };
};


export const ProjectsPage = () => {
  const { token, sheetInfoCache, opsConfig, setOpsConfig, combinedOpsAndData, setCombinedOpsAndData } = useAuth();  
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
  const { logMessages, doLog } = useLogger(60000);
  const [showLogPanel, setShowLogPanel] = useState(false);
  const [cachedToken, setCachedToken] = useState<{ token: LoginResponse; timestamp: number } | null>(null);
  
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
  async function fetchData() {
    console.log('useEffect run');
    if (token) {
      const ops = await getSheetOps({ token }, sheetInfoCache);
      if (opsConfig &&combinedOpsAndData) {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects only...' }));
        const res = await loadMainData(ops);
        updateDoneParentIds(combinedOpsAndData, res.operationList, logParam);
        prepareData(res, opsConfig);
        setIsLoading(prev=>({ ...prev, listLoading: '' }));
      } else {
        setIsLoading(prev => ({ ...prev, listLoading: 'Loading projects and config...' }));
      
        
        await getOpsAndMainList(ops, logParam).then(async res => {
          setOpsConfig(res); // Save the complete response (including functions)
          //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
          const combined = await combineOpsConfigWithFreedCampData(res, freeCampOpsWithCache, logParam);          

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
          prepareData(res, res);
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

    function prepareData(dataRes: { operationList: IOperationWithLineNumber[];}, res: IOpsConfig) {
      const list = dataRes.operationList.map((item, index) => ({ ...item, line: index + 2 })).filter(item => {
        return res.groupAndMainProjectMapping.actions.reduce((acc, action) => {
          return acc || item[getTaskIdColumnName(action)] != 'done';
        }, false) && item.文件.trim() !== '';
      });
      
      setProjectList(list);
    }
  }
  useEffect(() => {
    if (token && !startupRunOnce.current) {
      startupRunOnce.current = true;
      fetchData();
    }
  }, [token]);
  


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
    deleteItemActionTask,
    completeDateUpdater,
    doLog,
    fetchData,
    setErrorDialog,
  };
  return (
    <>
      <ErrorDialog 
        show={errorDialog.show}
        message={errorDialog.message}
        onClose={() => setErrorDialog({ show: false, message: '' })}
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

      <h1>Enyu Site</h1>
      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '2rem' }}>
        <thead>
          <tr>
            <th><button className="btn btn-create" onClick={fetchData}>Reload</button></th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文件</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文章名</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>作者</th>
            {opsConfig?.groupAndMainProjectMapping.actions.map((action) => (
              <th key={action} style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>
                {action}
              </th>
            ))}
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>Line</th>
          </tr>
        </thead>
        <tbody>
          {
            projectList.map((p,idx) => {
              return <tr key={p.文章名+idx}>
                <td><button className="btn btn-create" onClick={
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
                          await processOperation(ops, freeCampOpsWithCache, combinedOpsAndData, p, logs);
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
                }>Create</button></td>
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
                  { renderSyncActionCell(p, actionCellDeps, 'Sync from FreedCamp') }
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
