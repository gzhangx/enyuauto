import { useEffect, useState, useRef, useCallback } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import { getOpsAndMainList, getTaskIdColumnName, processOperation, type OperationWithDueDates } from '../lib/main_ops';
import { ErrorDialog } from './ErrorDialog';


type OperationWithDueDatesWithLineNumber = OperationWithDueDates & { line: number };

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
  const { token } = useAuth();  
  const [projectList, setProjectList] = useState<OperationWithDueDatesWithLineNumber[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  const [isLoading, setIsLoading] = useState('');
  const [errorDialog, setErrorDialog] = useState<{ show: boolean; message: string }>({ show: false, message: '' });
  const opsDataRef = useRef<Awaited<ReturnType<typeof getOpsAndMainList>> | null>(null);
  const { logMessages, doLog, animationDuration } = useLogger(5000);
  
  useEffect(() => {
    console.log('useEffect run');
    if (token) {
      setIsLoading('Loading projects...');
      getOpsAndMainList(token, { doLog }).then(res => {
        opsDataRef.current = res; // Save the complete response (including functions)
        //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
        const list = res.operationList.map((item, index) => ({ ...item, line: index + 2 })).filter(item => {
          return res.groupAndMainProjectMapping.actions.reduce((acc, action) => {
            return acc || item[getTaskIdColumnName(action)] != 'done';
          }, false) && item.文件.trim() !== '';
        });
        setProjectList(list);
      }).catch(error => {
        console.error('Error loading projects:', error);
        setResponseData(`Error: ${error.message || String(error)}`);
        doLog(`Error: ${error.message || String(error)}`);
        setErrorDialog({ show: true, message: `Failed to load projects:\n${error.message || String(error)}` });
      }).finally(() => {
        setIsLoading('');
      });
    }
  },[token, doLog]);

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
            <th></th>
            <th></th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文件</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文章名</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>作者</th>            
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>Line</th>
          </tr>
        </thead>
        <tbody>
          {
            projectList.map((p,idx) => {
              return <tr key={p.文章名+idx}>
                <td><button className="btn btn-create" onClick={
                  async () => {
                    if (opsDataRef.current === null) {
                      setResponseData('Operation data not loaded yet.');
                      return;
                    }
                    await processOperation(opsDataRef.current, p.line, { doLog });
                    //const retData = await createOrDelProject(p.line, 'main');
                    //setResponseData(JSON.stringify(retData, null, 2));
                  }
                }>Create</button></td>
                <td><button className="btn btn-delete" onClick={async () => {
                    //const retData = await createOrDelProject(p.line, 'del');
                    //setResponseData(JSON.stringify(retData, null, 2));
                }}>Delete</button></td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>
                  <a href={p.文件} target="_blank">{p.文件}</a>
                </td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.文章名}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.作者}</td>                
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.line}</td>                
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
