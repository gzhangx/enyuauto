
import type { ProjectTaskParams, CreateTaskResponse, FreedCampOps, FreedCampProcessor, ICurrentSessionData, LoginParams, LoginResponse } from '../../shared/freedcampTypes';
import { freedcampApi } from './api';




async function login({ username, password }: LoginParams): Promise<LoginResponse> {    
    const res = await freedcampApi({
        subAction: 'login',
            username: username || '',
            password: password || '',
    });

    return res as unknown as LoginResponse;
}

function getProcessor(loginToken: LoginResponse): FreedCampProcessor {
    // Initialize session by logging in

    async function doPostAttachment(taskId: string, params: ProjectTaskParams): Promise<any> {
        const res = await freedcampApi({
            subAction: 'doPostAttachment',            
                cookies: loginToken.Cookie,
                taskId,
                description: params.description,
                assignedToId: params.assigned_to_id,
                dueDate: params.due_date,
        });
        return res;
    }

    async function createTask(projectAndGroup: ProjectTaskParams, title: string): Promise<CreateTaskResponse> {
        const res = await freedcampApi<CreateTaskResponse>({
            subAction: 'createTask',
                cookies: loginToken.Cookie,
                projectId: projectAndGroup.project_id,
                title,
                description: projectAndGroup.description,
                dueDate: undefined,
                assignedToId: undefined,
                parentId: projectAndGroup.h_parent_id,
        });
        return res;
    }

    async function deleteTask(taskId: string | number): Promise<any> {
        const res = await freedcampApi({
            subAction: 'deleteTask',
            cookies: loginToken.Cookie,
            taskId: String(taskId),
        });
        return res;
    }

    async function getSessionCurrentData(): Promise<ICurrentSessionData> {
        const res = await freedcampApi({            
            subAction: 'getSessionCurrentData',            
            cookies: loginToken.Cookie,
            
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

export const freedCampOps: FreedCampOps = {
    login,
    getProcessor,
};

