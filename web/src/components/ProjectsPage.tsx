import { useEffect, useState, useRef, useCallback } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import {
  getFreeCampAndUpdateOperations, getOpsAndMainList, getSheetOps, processOperation, deleteItemActionTask,
  getActionsToPerform,
  type FreeCampAndUpdateOperations,
  type ActionsToPerformInfo,
  loadMainData,
  combineOpsConfigWithFreedCampData,
  type DebugLog,
  formatLocalDateYyyyMmDd,
  completeDateUpdater,  
  updateDoneParentIds,
  updateDoneColumn,
  type IOneDriveDirInfo,
  createMsExcelDataOps,
} from '../../shared/main_ops';
import type { ActionType } from '../../shared/types';
import { ErrorDialog } from './ErrorDialog';
import { DoGoogleSignIn, freedCampOps } from '../lib/util';
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
      console.log('show msg in doLog',msg);
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
  const { token, msToken, msAccount, msLoginRedirect, msLogout, sheetInfoCache, opsConfig, setOpsConfig, combinedOpsAndData, setCombinedOpsAndData, freedCampCredentials, useMsOps, setUseMsOps, authLoadingStatus } = useAuth();  
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
  type CreateTasksDialogState = {
    operation: IOperationWithLineNumberAndParentTaskId;
    oneDriveFolders: IOneDriveDirInfo[];
    actions: ActionsToPerformInfo[];
    selected: Partial<Record<ActionType, boolean>>;
    useEnglishTemplate: Partial<Record<ActionType, boolean>>;
    useAITemplate: Partial<Record<ActionType, boolean>>;
    selectedEditor?: Partial<Record<ActionType, string>>;
  };
  const [createTasksDialog, setCreateTasksDialog] = useState<CreateTasksDialogState | null>(null);
  const sheetOpsRef = useRef<Awaited<ReturnType<typeof getSheetOps>> | null>(null);
  const { logMessages, doLog, criticalError, closeCriticalError } = useLogger(60000);
  const [showLogPanel, setShowLogPanel] = useState(false);
  const [cachedToken, setCachedToken] = useState<{ token: LoginResponse; timestamp: number } | null>(null);
  type OwnerInfoEntry = { article: string; name: string; assignedOn: string; status: string; taskId?: string; timestamp: number; };
  const [showAllProjects, setShowAllProjects] = useState(false);
  const [daysShowAlertAfterComplete, setDaysShowAlertAfterComplete] = useState(3);
  const [doneColumnSavingLine, setDoneColumnSavingLine] = useState<number | null>(null);
  const [ownerInfoDialog, setOwnerInfoDialog] = useState<{ title: string; entries: OwnerInfoEntry[] } | null>(null);

  const originalFreedCampOps = getFreeCampAndUpdateOperations(freedCampOps);

  const closeOwnerInfoDialog = () => setOwnerInfoDialog(null);
  const getOwnerInfoEntriesForAction = (action: ActionType): OwnerInfoEntry[] => {
    const byName = new Map<string, OwnerInfoEntry>();
    for (const p of fullProjectList) {
      const freedCampItem = p[`${action} FreeCamp Item`];
      if (!freedCampItem || !freedCampItem.assigned_to_fullname) continue;
      const timestamp = freedCampItem.completed_ts ?? 0;
      const entry: OwnerInfoEntry = {
        article: p.文章名,
        name: freedCampItem.assigned_to_fullname,
        assignedOn: timestamp ? formatLocalDateYyyyMmDd(timestamp) : 'Unknown',
        status: freedCampItem.status_title || 'Unknown',
        taskId: freedCampItem.id,
        timestamp,
      };
      const existing = byName.get(entry.name);
      if (!existing || entry.timestamp > existing.timestamp) {
        byName.set(entry.name, entry);
      }
    }
    return Array.from(byName.values()).sort((a, b) => b.timestamp - a.timestamp || a.name.localeCompare(b.name));
  };
  const openOwnerInfoDialogForAction = (action: ActionType) => {
    setOwnerInfoDialog({
      title: `Owner assignment for ${action}`,
      entries: getOwnerInfoEntriesForAction(action),
    });
  };
  

  const startupRunOnce = useRef(false);
  const freeCampOpsWithCache: FreeCampAndUpdateOperations = {
    ...originalFreedCampOps,
    getFreedCampToken: async (opsAndTemplates, forceReload) => {
      const now = Date.now();
      const oneHour = 60 * 60 * 1000; // 1 hour in milliseconds
      
      // Check if cached token exists and is still valid
      if (!forceReload && cachedToken && (now - cachedToken.timestamp) < oneHour) {
        doLog('Using cached FreedCamp token');
        return cachedToken.token;
      }
      if (forceReload) {
        doLog('Force reload requested for FreedCamp token');
      }
      
      // Get new token if cache is expired or doesn't exist
      doLog('Fetching new FreedCamp token');
      const newToken = await originalFreedCampOps.getFreedCampToken(opsAndTemplates, forceReload);
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
          const combined = await combineOpsConfigWithFreedCampData(res, freeCampOpsWithCache, freedCampCredentials, logParam, fullReload);          

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
  if (authLoadingStatus || isLoading.listLoading || isLoading.freeCampLoading || isLoading.projectButtonAction) {
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
        {authLoadingStatus && <h2 style={{ margin: 0, color: '#555' }}>{authLoadingStatus}</h2>}
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
    msToken,
    wpToken: freedCampCredentials?.enyu_wp_token,
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

      {createTasksDialog && combinedOpsAndData && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 2100,
          }}
          onClick={() => setCreateTasksDialog(null)}
        >
          <div
            style={{
              backgroundColor: 'white',
              padding: '1.5rem',
              borderRadius: '8px',
              maxWidth: '520px',
              width: '92%',
              maxHeight: '80vh',
              overflow: 'auto',
              boxShadow: '0 4px 20px rgba(0,0,0,0.3)',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <h2 style={{ margin: '0 0 0.75rem 0', fontSize: '1.15rem' }}>
              Create FreedCamp tasks — {createTasksDialog.operation['文件']}
            </h2>            
            <ul style={{ listStyle: 'none', padding: 0, margin: '0 0 1.25rem 0' }}>
              {createTasksDialog.actions.map((row) => (
                <li
                  key={row.action}
                  style={{
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '8px',
                    padding: '8px 0',
                    borderBottom: '1px solid #eee',
                    fontSize: '14px',
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <input
                      type="checkbox"
                      id={`create-task-${row.action}`}
                      checked={createTasksDialog.selected[row.action] === true}
                      onChange={(e) => {
                        const checked = e.target.checked;
                        setCreateTasksDialog((prev) => {
                          if (!prev) return prev;
                          //const mapping =
                          //  combinedOpsAndData.opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId;
                          const selected = { ...prev.selected };

                          // const collectDescendants = (par: ActionType): ActionType[] => {
                          //   const out: ActionType[] = [];
                          //   for (const item of prev.actions) {
                          //     const cfg = mapping[item.action];
                          //     if (cfg?.subTaskOf === par) {
                          //       out.push(item.action, ...collectDescendants(item.action));
                          //     }
                          //   }
                          //   return out;
                          // };

                          if (checked) {
                            selected[row.action] = true;
                            // let parent = mapping[row.action]?.subTaskOf;
                            // while (parent) {
                            //   if (prev.actions.some((x) => x.action === parent)) {
                            //     selected[parent] = true;
                            //   }
                            //   parent = mapping[parent]?.subTaskOf;
                            // }
                          } else {
                            selected[row.action] = false;
                            // for (const desc of collectDescendants(row.action)) {
                            //   selected[desc] = false;
                            // }
                          }
                          return { ...prev, selected };
                        });
                      }}
                    />
                    <label htmlFor={`create-task-${row.action}`} style={{ cursor: 'pointer', flex: 1 }}>
                      <strong>{row.action}</strong>
                      <span 
                        style={{ color: '#666', marginLeft: '8px' }} 
                        title={`Due: ${createTasksDialog.operation[`${row.action} Due Date` as keyof IOperationWithLineNumberAndParentTaskId] as string || 'N/A'}`}
                      >
                        {row.editorName}
                      </span>
                    </label>
                  </div>
                  {/* Editor dropdown: show full name from editorInfoMap, default to row.editorName or sheet value */}
                  {createTasksDialog.selected[row.action] &&
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginLeft: '20px' }}>
                      <select
                        value={createTasksDialog.selectedEditor?.[row.action] ?? ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setCreateTasksDialog((prev) => {
                            if (!prev) return prev;
                            const selectedEditor = { ...(prev.selectedEditor || {}) };
                            selectedEditor[row.action] = val;
                            return { ...prev, selectedEditor };
                          });
                        }}
                        style={{ marginLeft: '8px', fontSize: '13px', padding: '4px' }}
                      >
                        <option value="">(select editor)</option>
                        {Object.entries(combinedOpsAndData.opsConfig.editorInfoMap).map(([shortName, info]) => {
                          const pretty = (() => {
                            const prettyName = info.print_name || info.shortName || '';
                            const normalizedTitle = (info.title || '').toLowerCase();
                            const isEng = normalizedTitle === 'brother' || normalizedTitle === 'sister';
                            return isEng ? `${info.title} ${prettyName}` : `${prettyName}${info.title || ''}`;
                          })();
                          return <option key={shortName} value={shortName}>{info.print_name || pretty}</option>;
                        })}
                      </select>
                    </div>
                  }
                  {row.hasEnglishTemplate && createTasksDialog.selected[row.action] && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginLeft: '20px' }}>
                      <input
                        type="checkbox"
                        id={`use-english-${row.action}`}
                        checked={createTasksDialog.useEnglishTemplate[row.action] === true}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setCreateTasksDialog((prev) => {
                            if (!prev) return prev;
                            const useEnglishTemplate = { ...prev.useEnglishTemplate };
                            useEnglishTemplate[row.action] = checked;
                            return { ...prev, useEnglishTemplate };
                          });
                        }}
                      />
                      <label htmlFor={`use-english-${row.action}`} style={{ cursor: 'pointer', color: '#555', fontSize: '13px' }}>
                        Use English template
                      </label>                      
                    </div>
                  )}
                  {row.hasAITemplate && createTasksDialog.selected[row.action] && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginLeft: '20px' }}>
                      <input
                        type="checkbox"
                        id={`use-ai-${row.action}`}
                        checked={createTasksDialog.useAITemplate[row.action] !== false}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setCreateTasksDialog((prev) => {
                            if (!prev) return prev;
                            const useAITemplate = { ...prev.useAITemplate };
                            useAITemplate[row.action] = checked;
                            return { ...prev, useAITemplate };
                          });
                        }}
                      />
                      <label htmlFor={`use-ai-${row.action}`} style={{ cursor: 'pointer', color: '#555', fontSize: '13px' }}>
                        Use AI template
                      </label>
                    </div>
                  )}
                </li>
              ))}
            </ul>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button
                type="button"
                onClick={() => setCreateTasksDialog(null)}
                style={{
                  padding: '8px 18px',
                  borderRadius: '4px',
                  border: '1px solid #ccc',
                  background: '#f5f5f5',
                  cursor: 'pointer',
                  fontSize: '14px',
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="btn btn-create"
                onClick={async () => {
                  const dlg = createTasksDialog;
                  if (!dlg || !combinedOpsAndData) return;
                  const selectedSet = new Set<ActionType>(
                    dlg.actions.filter((a) => dlg.selected[a.action]).map((a) => a.action),
                  );
                  if (selectedSet.size === 0) {
                    setErrorDialog({
                      show: true,
                      message: 'Select at least one action to create, or press Cancel.',
                    });
                    return;
                  }
                  // apply selected editor choices (shortName) into the operation before creating tasks
                  if (dlg.selectedEditor) {
                    for (const k of Object.keys(dlg.selectedEditor) as ActionType[]) {
                      const v = dlg.selectedEditor[k];
                      if (v) (dlg.operation as any)[k] = v;
                    }
                  }
                  setCreateTasksDialog(null);
                  const ops = sheetOpsRef.current;
                  if (!ops) return;
                  setIsLoading((prev) => ({ ...prev, projectButtonAction: 'Creating project...' }));
                  const logs: DebugLog = {
                    doLog: (msg) => {
                      console.log('doLog',msg);
                      setProgressText(msg);
                    },
                  };
                  try {
                    const useAITemplateMap = new Map<ActionType, boolean>();
                    for (const action of Object.keys(dlg.useAITemplate) as ActionType[]) {
                      if (dlg.useAITemplate[action] !== false) {
                        useAITemplateMap.set(action, true);
                      }
                    }
                    const useEnglishTemplateMap = new Map<ActionType, boolean>();
                    for (const action of Object.keys(dlg.useEnglishTemplate) as ActionType[]) {
                      if (dlg.useEnglishTemplate[action]) {
                        useEnglishTemplateMap.set(action, true);
                      }
                    }
                    await processOperation(
                      ops,
                      freeCampOpsWithCache,
                      combinedOpsAndData,
                      dlg.operation,
                      dlg.oneDriveFolders,
                      logs,
                      '',
                      selectedSet,
                      useAITemplateMap,
                      useEnglishTemplateMap,
                      msToken || undefined,
                    );
                  } catch (error: unknown) {
                    const err = error as { message?: string };
                    console.error('Error creating project:', error);
                    setErrorDialog({
                      show: true,
                      message: `Failed to create project:\n${err.message || String(error)}`,
                    });
                  } finally {
                    setIsLoading((prev) => ({ ...prev, projectButtonAction: '' }));
                    setProgressText('');
                  }
                }}
              >
                Create selected
              </button>
            </div>
          </div>
        </div>
      )}

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
        <select
          value={useMsOps ? 'ms' : 'google'}
          disabled = {!DoGoogleSignIn}
          onChange={e => {
            setUseMsOps(e.target.value === 'ms');
            startupRunOnce.current = false;
          }}
          style={{ fontSize: '14px', padding: '2px 6px', borderRadius: '4px', border: '1px solid #ccc', cursor: 'pointer' }}
        >
          <option value="google">Google Sheets</option>
          <option value="ms">Microsoft Excel (SharePoint)</option>
        </select>
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
                <button
                  type="button"
                  onClick={() => openOwnerInfoDialogForAction(action)}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '0.35rem',
                    border: 'none',
                    background: 'none',
                    padding: 0,
                    margin: 0,
                    font: 'inherit',
                    color: 'inherit',
                    cursor: 'pointer',
                  }}
                  title={`Show all owners for ${action} sorted by date descending`}
                >
                  <span>{action}</span>
                  <span style={{ fontSize: '0.9em' }}>👤</span>
                </button>
              </th>
            ))}
            <th
              style={stickyHeaderCellStyle}
              title='Sets the sheet column "done" to Y so this row is hidden unless "Show all" is checked.'
            >
              Done
            </th>
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
                    const ops = sheetOpsRef.current;
                    if (!ops || !combinedOpsAndData) return;
                    setIsLoading(prev => ({ ...prev, projectButtonAction: 'Preparing…' }));
                    try {
                      let oneDriveFolders: IOneDriveDirInfo[] = [];
                      if (p.mainFolder) {
                        if (!msToken) {
                          setErrorDialog({ show: true, message: 'Microsoft token not available. Please sign in to Microsoft to access OneDrive folders.' });
                          return;
                        }
                        oneDriveFolders = await fetchOneDriveChildren(msToken, p.mainFolder);
                      }
                      const previewLog: DebugLog = { doLog: (msg) => console.log(msg) };
                      const previewOp = { ...p } as IOperationWithLineNumberAndParentTaskId;
                      const actions = getActionsToPerform(previewOp, combinedOpsAndData, p['文件'], previewLog);
                      if (actions.length === 0) {
                        setErrorDialog({
                          show: true,
                          message:
                            'No new FreedCamp tasks to create for this row (all actions already have tasks or are excluded).',
                        });
                        return;
                      }
                      const selected: Partial<Record<ActionType, boolean>> = {};
                      const useEnglishTemplate: Partial<Record<ActionType, boolean>> = {};
                      const useAITemplate: Partial<Record<ActionType, boolean>> = {};
                      const selectedEditor: Partial<Record<ActionType, string>> = {};
                      for (const a of actions) {
                        if (a.hasAITemplate) {
                          useAITemplate[a.action] = true;
                        }
                        // default selected editor to the value in the sheet row if present
                        // `p` is the operation row; it may contain the short editor key
                        // otherwise leave undefined
                        const maybeEditor = (p as any)[a.action] as string | undefined;
                        if (maybeEditor) {
                          selectedEditor[a.action] = maybeEditor;
                        } else if (a.editorName && combinedOpsAndData?.opsConfig?.editorInfoMap) {
                          // try to find a matching shortName by comparing pretty display name
                          const entries = Object.entries(combinedOpsAndData.opsConfig.editorInfoMap);
                          for (const [shortName, info] of entries) {
                            const prettyName = info.print_name || info.shortName || '';
                            const normalizedTitle = (info.title || '').toLowerCase();
                            const isEng = normalizedTitle === 'brother' || normalizedTitle === 'sister';
                            const display = isEng ? `${info.title} ${prettyName}` : `${prettyName}${info.title || ''}`;
                            if (display === a.editorName) {
                              selectedEditor[a.action] = shortName;
                              break;
                            }
                          }
                        }
                      }
                      //for (const a of actions) selected[a.action] = true;
                      setCreateTasksDialog({ operation: p, oneDriveFolders, actions, selected, useEnglishTemplate, useAITemplate, selectedEditor });
                    } catch (error: unknown) {
                      const err = error as { message?: string };
                      console.error('Error preparing project creation:', error);
                      setErrorDialog({
                        show: true,
                        message: `Failed to prepare project creation:\n${err.message || String(error)}`,
                      });
                    } finally {
                      setIsLoading(prev => ({ ...prev, projectButtonAction: '' }));
                    }
                  }
                }>Create</button>}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>
                  {p.mainFolder
                    ? <a href={`${p.mainFolder.replace(/\/$/, '')}/${p.文件}`} target="_blank">{p.文件}</a>
                    : p.文件}
                </td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.文章名}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.作者}</td>
                {opsConfig?.groupAndMainProjectMapping.actions.map((action) => (
                  <td key={action} style={{ border: '1px solid #ddd', padding: '12px' }}>
                    {renderActionCell(p, action, actionCellDeps)}
                  </td>
                ))}
                <td style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'center' }}>
                  <input
                    type="checkbox"
                    title='Maps to sheet column "done": Y hides this row from the default list.'
                    checked={(p.done || '').toUpperCase() === 'Y'}
                    disabled={
                      !opsConfig ||
                      opsConfig.headers.indexOf('done') < 0 ||
                      doneColumnSavingLine === p.itemPositionOnSheet
                    }
                    onChange={async (e) => {
                      const checked = e.target.checked;
                      const ops = sheetOpsRef.current;
                      if (!ops || !opsConfig) return;
                      setDoneColumnSavingLine(p.itemPositionOnSheet);
                      try {
                        await updateDoneColumn(ops, opsConfig, p, checked, { doLog });
                        setFullProjectList((prev) =>
                          prev.map((row) =>
                            row.itemPositionOnSheet === p.itemPositionOnSheet
                              ? { ...row, done: checked ? 'Y' : '', isFinished: checked }
                              : row,
                          ),
                        );
                      } catch (error: unknown) {
                        const err = error as { message?: string };
                        setErrorDialog({
                          show: true,
                          message: err.message || String(error),
                        });
                      } finally {
                        setDoneColumnSavingLine(null);
                      }
                    }}
                  />
                </td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.itemPositionOnSheet}
                  { renderSyncActionCell(p, freedCampCredentials, actionCellDeps, 'Sync from FreedCamp') }
                </td>
              </tr>
            })
          }
        </tbody>
      </table>      
      {ownerInfoDialog && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            backgroundColor: 'rgba(0,0,0,0.35)',
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            zIndex: 2000,
            padding: '1rem',
          }}
          onClick={closeOwnerInfoDialog}
        >
          <div
            style={{
              backgroundColor: 'white',
              borderRadius: '8px',
              boxShadow: '0 12px 30px rgba(0,0,0,0.25)',
              maxWidth: '600px',
              width: '100%',
              maxHeight: '90vh',
              overflow: 'auto',
              padding: '1rem 1.25rem',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '1rem' }}>
              <h3 style={{ margin: 0, fontSize: '1.1rem' }}>{ownerInfoDialog.title}</h3>
              <button
                type="button"
                onClick={closeOwnerInfoDialog}
                style={{ border: 'none', background: '#6c757d', color: 'white', borderRadius: '4px', padding: '4px 10px', cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
            <div style={{ marginTop: '1rem' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Article</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Owner</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Assigned</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Status</th>
                    <th style={{ textAlign: 'left', padding: '8px', borderBottom: '1px solid #ddd' }}>Task ID</th>
                  </tr>
                </thead>
                <tbody>
                  {ownerInfoDialog.entries.length === 0 ? (
                    <tr>
                      <td colSpan={5} style={{ padding: '12px', borderBottom: '1px solid #eee', textAlign: 'center', color: '#555' }}>
                        No owner assignment details available.
                      </td>
                    </tr>
                  ) : (
                    ownerInfoDialog.entries.map((entry, index) => (
                      <tr key={`${entry.name}-${index}`}>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.article}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.name}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.assignedOn}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.status}</td>
                        <td style={{ padding: '8px', borderBottom: '1px solid #eee' }}>{entry.taskId || '—'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
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
