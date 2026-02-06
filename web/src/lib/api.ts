
const actions = ['校对','二校', '美编', '发布'] as const;
export type ActionType = typeof actions[number];
export function getActions(): readonly ActionType[] {
    return actions ;
}

const enyuApiBaseUrl = 'https://fk9u03bqm3.execute-api.us-east-1.amazonaws.com/default/enyu_auto';
export type ProjectItem = {
        '文件': string;
        '作者': string;
    '文章名': string;
    line: number;
} & {
        [key in ActionType as `${key} TaskId`]: string;
};

export async function getProjectList() {
    const response = await fetch(enyuApiBaseUrl + '?line=14&doit1=true&del=true&action=getList');
    const projects = await response.json(); 
    console.log('projects', projects);
    const items = (projects.operationList as any as ProjectItem[]).map((item, idx) => {
        item.line = idx + 2;
        return item;
    });      
    
    return items;
}


export async function createOrDelProject(line: number, action: 'main' | 'del' = 'main') {
    const response = await fetch(enyuApiBaseUrl + '?action='+action+'&line='+line);
    const projects = await response.json(); 
    console.log('projects', projects);
    return projects;      
}

// Freedcamp API types
type FreedcampLoginParams = {
    username: string;
    password: string;
};

type FreedcampCreateTaskParams = {
    projectId: string;
    title: string;
    description?: string;
    assignedToId?: string;
    dueDate?: string;
    parentId?: string;
};

type FreedcampDeleteTaskParams = {
    taskId: string;
};

type FreedcampDoPostAttachmentParams = {
    taskId: string;
    description?: string;
    assignedToId?: string;
    dueDate?: string;
};

type FreedcampParams = 
    | { subAction: 'login'; params: FreedcampLoginParams }
    | { subAction: 'getSessionCurrentData'; params?: never }
    | { subAction: 'createTask'; params: FreedcampCreateTaskParams }
    | { subAction: 'deleteTask'; params: FreedcampDeleteTaskParams }
    | { subAction: 'doPostAttachment'; params: FreedcampDoPostAttachmentParams };

export async function freedcampApi<T = any>(request: FreedcampParams): Promise<T> {
    const body = {
        action: 'freedcamp',
        subAction: request.subAction,
        ...request.params
    };

    const response = await fetch(enyuApiBaseUrl, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(body),
    });

    if (!response.ok) {
        throw new Error(`Freedcamp API error: ${response.statusText}`);
    }

    return await response.json();
}

