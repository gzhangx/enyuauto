import { useEffect, useState } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import { getOpsAndMainList, getTaskIdColumnName, type OperationWithDueDates } from '../lib/main_ops';


type OperationWithDueDatesWithLineNumber = OperationWithDueDates & { line: number };

type LogMessage = {
  id: number;
  text: string;
  timestamp: number;
};

export const ProjectsPage = () => {
  const { token } = useAuth();  
  const [projectList, setProjectList] = useState<OperationWithDueDatesWithLineNumber[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  const [logMessages, setLogMessages] = useState<LogMessage[]>([]);
  const [logIdCounter, setLogIdCounter] = useState(0);
  
  useEffect(() => {
    console.log('useEffect run');
    if (token) {
      getOpsAndMainList(token, {
        doLog: (msg: string) => {
          console.log(msg);
          const newLog: LogMessage = {
            id: logIdCounter,
            text: msg,
            timestamp: Date.now()
          };
          setLogIdCounter(prev => prev + 1);
          setLogMessages(prev => [...prev, newLog]);
          
          // Remove message after 5 seconds
          setTimeout(() => {
            setLogMessages(prev => prev.filter(log => log.id !== newLog.id));
          }, 5000);
        }
      }).then(res => { 
        //const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } = res;
        const list = res.operationList.map((item, index) => ({ ...item, line: index + 2 })).filter(item => {
          return res.groupAndMainProjectMapping.actions.reduce((acc, action) => {            
            return acc || item[getTaskIdColumnName(action)] != 'done';
          }, false) && item.文件.trim() !== '';
        });
        setProjectList(list);
        setResponseData('');
      });
    }
  },[token]);

  // Calculate animation speed based on number of messages
  const animationDuration = logMessages.length > 10 ? 0.3 : logMessages.length > 5 ? 0.5 : 1;

  return (
    <>
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
