import { useEffect, useState, useRef, useCallback, use } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import {
  getFreeCampAndUpdateOperations, getOpsAndMainList, getSheetOps, processOperation,deleteItemActionTask,
  type FreeCampAndUpdateOperations,
  loadMainData,
  combineOpsConfigWithFreedCampData,
  type DebugLog,
  type ICombinedOpsAndFreeCampData
} from '../../shared/main_ops';
import { ErrorDialog } from './ErrorDialog';
import { freedCampOps } from '../lib/util';
import type { LoginResponse } from '../../shared/freedcampTypes';
import { getTaskIdColumnName, type IOperationWithLineNumber, type IOpsConfig } from '../../shared/opsTypes';
import type { ActionType } from '../lib/api';
import { set } from '@gzhangx/googleapi/lib/util';


type LogMessage = {
  id: number;
  text: string;
  timestamp: number;
};

const useLogger = (displayDuration = 5000) => {
  const [logMessages, setLogMessages] = useState<LogMessage[]>([]);
  const logIdCounterRef = useRef(0);
  
  const doLog = useCallback((msg: string) => {
    console.log(msg);
    const newLog: LogMessage = {
      id: logIdCounterRef.current++,
      text: msg,
      timestamp: Date.now()
    };
    setLogMessages(prev => [...prev, newLog]);
    
    setTimeout(() => {
      setLogMessages(prev => prev.filter(log => log.id !== newLog.id));
    }, displayDuration);
  }, [displayDuration]);

  const animationDuration = logMessages.length > 10 ? 0.3 : logMessages.length > 5 ? 0.5 : 1;

  return { logMessages, doLog, animationDuration };
};

export const ProjectsPage = () => {
  const { token, sheetInfoCache, opsConfig, setOpsConfig, combinedOpsAndData, setCombinedOpsAndData } = useAuth();  
  const [projectList, setProjectList] = useState<IOperationWithLineNumber[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  const [isLoading, setIsLoading] = useState('');
  const [progressText, setProgressText] = useState('');
  const [errorDialog, setErrorDialog] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const sheetOpsRef = useRef<Awaited<ReturnType<typeof getSheetOps>> | null>(null);
  const { logMessages, doLog, animationDuration } = useLogger(5000);
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
  async function fetchData() {
    console.log('useEffect run');
    if (token) {
      const ops = await getSheetOps({ token }, sheetInfoCache);
      if (opsConfig) {
        setIsLoading('Loading projects only...');
        const res = await loadMainData(ops);
        await prepareData(res, opsConfig);
        setIsLoading('');
      } else {
        setIsLoading('Loading projects and config...');
      
        const logs = { getOperations: () => [], doLog }
        await getOpsAndMainList(ops, logs).then(async res => {
          setOpsConfig(res); // Save the complete response (including functions)
          //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
          const combined = await combineOpsConfigWithFreedCampData(res, freeCampOpsWithCache, logs);
          setCombinedOpsAndData(combined);
          await prepareData(res, res);
        }).catch(error => {
          console.error('Error loading projects:', error);
          setResponseData(`Error: ${error.message || String(error)}`);
          doLog(`Error: ${error.message || String(error)}`);
          setErrorDialog({ show: true, message: `Failed to load projects:\n${error.message || String(error)}` });
        }).finally(() => {
          setIsLoading('');
        });
      }
      sheetOpsRef.current = ops;
    }

    async function prepareData(dataRes: { operationList: IOperationWithLineNumber[];}, res: IOpsConfig) {
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
  
  useEffect(() => {
    if (!opsConfig) return;
    console.log('Checking for done sub-tasks to mark main tasks as done...');
      opsConfig.groupAndMainProjectMapping.actions.forEach(action => {
        const actionInfo = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
        if (actionInfo && actionInfo.subTaskOf) {
          const subTaskOfAction = actionInfo.subTaskOf;
          projectList.forEach(item => {
            if (item[getTaskIdColumnName(subTaskOfAction)] === 'done' && !item[getTaskIdColumnName(action)]) {
              console.log(`EARNNN ${item.文件} ${item.文章名} ${action} as done for line ${item.itemPositionOnSheet} because ${subTaskOfAction} is done`);
            }
          });
        }
      });
  }, [opsConfig?.groupAndMainProjectMapping?.shortProjectNameToProjectId, combinedOpsAndData]);


  const renderActionCell = (p: IOperationWithLineNumber, action: ActionType) => {
    const taskId = p[getTaskIdColumnName(action)];
    
    if (!taskId) {
      return null;
    }

    if (taskId === 'done') {
      return <span style={{ color: 'green', fontWeight: 'bold' }}>Done</span>;
    }

    return (
      <button className="btn btn-delete" onClick={async () => {
        //const retData = await createOrDelProject(p.itemPositionOnSheet, 'del');
        //setResponseData(JSON.stringify(retData, null, 2));
        if (sheetOpsRef.current && opsConfig) {
          await deleteItemActionTask(sheetOpsRef.current, freeCampOpsWithCache, opsConfig, p, action, { getOperations: () => [], doLog });
        }
      }}>Delete {p[getTaskIdColumnName(action)]}</button>
    );
  };

  // Calculate animation speed based on number of messages
  if (isLoading) {
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
        <h2 style={{ margin: 0, color: '#555' }}>{isLoading}</h2>
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

  return (
    <>
      <ErrorDialog 
        show={errorDialog.show}
        message={errorDialog.message}
        onClose={() => setErrorDialog({ show: false, message: '' })}
      />

      {/* Scrolling Log Display */}
      <div style={{
        position: 'fixed',
        right: '20px',
        bottom: '20px',
        width: '300px',
        maxHeight: '80vh',
        overflow: 'hidden',
        pointerEvents: 'none',
        zIndex: 1000,
        display: 'flex',
        flexDirection: 'column-reverse',
        gap: '8px'
      }}>
        {logMessages.map((log) => {
          const age = Date.now() - log.timestamp;
          const isFadingOut = age > 4500;
          
          return (
            <div
              key={log.id}
              style={{
                backgroundColor: 'rgba(33, 150, 243, 0.9)',
                color: 'white',
                padding: '8px 12px',
                borderRadius: '4px',
                fontSize: '12px',
                wordWrap: 'break-word',
                boxShadow: '0 2px 4px rgba(0,0,0,0.2)',
                animation: `slideInFromBottom ${animationDuration}s ease-out`,
                opacity: isFadingOut ? 0 : 1,
                transition: `opacity 0.5s ease-out`,
                transformOrigin: 'bottom'
              }}
            >
              {log.text}
            </div>
          );
        })}
      </div>
      
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
                    setIsLoading('Creating project...');
                    try {
                      const ops = sheetOpsRef.current;
                      if (ops) {
                        const logs: DebugLog = {
                          getOperations: () => [],
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
                      setIsLoading('');
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
                    {renderActionCell(p, action)}
                  </td>
                ))}
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.itemPositionOnSheet}</td>
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
