import { useEffect, useState } from 'react'

import './App.css'
import type { ProjectItem } from './lib/api'

import { getProjectList } from './lib/api';
function App() {
  const [count, setCount] = useState(0)

  const [projectList, setProjectList] = useState<ProjectItem[]>([]);
  
  useEffect(() => {
    console.log('useEffect run');
    getProjectList().then(list => {
      list.reverse();
      setProjectList(list);
    });
  },['1']);
  return (
    <>
      <h1>Enyu Site</h1>
      <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '2rem' }}>
        <thead>
          <tr>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文章名</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>作者</th>
            <th style={{ border: '1px solid #ddd', padding: '12px', textAlign: 'left' }}>文件</th>
          </tr>
        </thead>
        <tbody>
          {
            projectList.map(p => {
              return <tr key={p.文章名}>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>
                  <a href={p.文件} target="_blank">{p.文件}</a>
                </td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.文章名}</td>
                <td style={{ border: '1px solid #ddd', padding: '12px' }}>{p.作者}</td>                
              </tr>
            })
          }
        </tbody>
      </table>
      <div className="card">
        <button onClick={() => setCount((count) => count + 1)}>
          count is {count}
        </button>
        <p>
          Edit <code>src/App.tsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">
        Click on the Vite and React logos to learn more
      </p>
    </>
  )
}

export default App
