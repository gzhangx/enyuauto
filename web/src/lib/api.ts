
const actions = ['校对','二校', '美编', '发布'] as const;
export type ActionType = typeof actions[number];
export function getActions(): readonly ActionType[] {
    return actions ;
}

export type ProjectItem = {
        '文件': string;
        '作者': string;
    '文章名': string;
    line: number;
} & {
        [key in ActionType as `${key} TaskId`]: string;
};

export async function getProjectList() {
    const response = await fetch('https://fk9u03bqm3.execute-api.us-east-1.amazonaws.com/default/enyu_auto?line=14&doit1=true&del=true&action=getList');
    const projects = await response.json(); 
    console.log('projects', projects);
    const items = (projects.operationList as any as ProjectItem[]).map((item, idx) => {
        item.line = idx + 2;
        return item;
    });      
    
    return items;
}


export async function createOrDelProject(line: number, action: 'main' | 'del' = 'main') {
    const response = await fetch('https://fk9u03bqm3.execute-api.us-east-1.amazonaws.com/default/enyu_auto?action='+action+'&line='+line);
    const projects = await response.json(); 
    console.log('projects', projects);
    return projects;      
}