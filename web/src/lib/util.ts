
import { freedcampApi } from './api';

interface LoginParams {
    username?: string;
    password?: string;
}

interface ProjectAndGroup {
    h_parent_id?: string;
    project_id: string;
    task_group_id: string;
    description?: string;
    "priority": 2;
}

interface CreateTaskResponse {
    result: any;
    id: number;
}

export interface TaskParams {
    h_parent_id?: string;
    description: string;
    assigned_to_id?: string;
    due_date?: string;
    "priority": 2;
}
interface Processor {
    doPostAttachment: (taskId: string, params: TaskParams) => Promise<any>;
    createTask: (projectAndGroup: ProjectAndGroup, title: string) => Promise<CreateTaskResponse>;
    deleteTask: (taskId: string | number) => Promise<any>;
    getSessionCurrentData(): Promise<ICurrentSessionData>;
}

export interface IUserInfo {
    user_id: string;
    full_name: string;
    email: string;
}

export interface IProjectInfo {
    project_id: string;
    //group_id: string;
    project_name: string; //文字二校 （Second Proofread)
    group_name: string; //EnYu_2026
}

export interface ICurrentSessionData {
    data: {
        users: IUserInfo[];
        projects: IProjectInfo[];
    }
}

async function login({ username, password }: LoginParams): Promise<LoginResponse> {    
    const res = await freedcampApi({
        subAction: 'login',
        params: {
            username: username || '',
            password: password || '',
        },
    });

    return res as unknown as LoginResponse;
}

function getProcessor(loginToken: LoginResponse): Processor {
    // Initialize session by logging in

    async function doPostAttachment(taskId: string, params: TaskParams): Promise<any> {
        const res = await freedcampApi({
            subAction: 'doPostAttachment',            
            params: {                
                cookies: loginToken.cookies,
                taskId,
                description: params.description,
                assignedToId: params.assigned_to_id,
                dueDate: params.due_date,
            },
        });
        return res;
    }

    async function createTask(projectAndGroup: ProjectAndGroup, title: string): Promise<CreateTaskResponse> {
        const res = await freedcampApi({
            subAction: 'createTask',
            params: {
                cookies: loginToken.cookies,
                projectId: projectAndGroup.project_id,
                title,
                description: projectAndGroup.description,
                dueDate: undefined,
                assignedToId: undefined,
                parentId: projectAndGroup.h_parent_id,
            },
        });
        return {
            result: res,
            id: res.data.tasks[0].id,
        };
    }

    async function deleteTask(taskId: string | number): Promise<any> {
        const res = await freedcampApi({
            subAction: 'deleteTask',
            params: {
                cookies: loginToken.cookies,
                taskId: String(taskId),
            },
        });
        return res;
    }

    async function getSessionCurrentData(): Promise<ICurrentSessionData> {
        const res = await freedcampApi({            
            subAction: 'getSessionCurrentData',
            params: {
                cookies: loginToken.cookies,
            },
        });
        return res as ICurrentSessionData;
    }

    return {
        doPostAttachment,
        createTask,
        deleteTask,
        getSessionCurrentData,
    };
}

export {
    login,
    getProcessor,
};
export type {
        LoginParams,
        LoginResponse,
        ProjectAndGroup,
        CreateTaskResponse,
        Processor
    };

