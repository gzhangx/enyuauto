

export interface ProjectItem {
        '文件': string;
        '作者': string;
        '文章名': string;
};

export async function getProjectList() {
    const response = await fetch('https://fk9u03bqm3.execute-api.us-east-1.amazonaws.com/default/enyu_auto?line=14&doit1=true&del=true&action=getList');
    const projects = await response.json(); 
    console.log('projects', projects);
    return projects.operationList as any as ProjectItem[];      
}