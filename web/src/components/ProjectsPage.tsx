import { useEffect, useState } from 'react'
import '../App.css'
import { useAuth } from '../contexts/AuthContext';
import { getOpsAndMainList, getTaskIdColumnName, type OperationWithDueDates } from '../lib/main_ops';


type OperationWithDueDatesWithLineNumber = OperationWithDueDates & { line: number };
export const ProjectsPage = () => {
  const { token } = useAuth();  
  const [projectList, setProjectList] = useState<OperationWithDueDatesWithLineNumber[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  
  useEffect(() => {
    console.log('useEffect run');
    if (token) {
      getOpsAndMainList(token, {
        doLog: (msg: string) => {
          console.log(msg);
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

  return (
    <>
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
