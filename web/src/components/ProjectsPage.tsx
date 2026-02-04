import { useEffect, useState } from 'react'
import '../App.css'
import type { ProjectItem, } from '../lib/api'
import { getProjectList, getActions, createOrDelProject } from '../lib/api';

export const ProjectsPage = () => {
  const [projectList, setProjectList] = useState<ProjectItem[]>([]);
  const [responseData, setResponseData] = useState<string>('');
  
  useEffect(() => {
    console.log('useEffect run');
    getProjectList().then(list => {
      list.reverse();
      list = list.filter(item => {
        let allIdDone = true;
        if (!item.文件) return false;
        for (const action of getActions()) {
          const taskId = item[`${action} TaskId`];
          console.log('taskId', taskId,`${action}TaskId`);
          if (taskId !== 'done') allIdDone = false;
        }
        console.log(item)
        return !allIdDone;
      })
      setProjectList(list);
    });
  },['1']);

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
                    const retData = await createOrDelProject(p.line, 'main');
                    setResponseData(JSON.stringify(retData, null, 2));
                  }
                }>Create</button></td>
                <td><button className="btn btn-delete" onClick={async () => {
                    const retData = await createOrDelProject(p.line, 'del');
                    setResponseData(JSON.stringify(retData, null, 2));
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
