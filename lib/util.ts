import * as gs from '@gzhangx/googleapi';
import * as secs from '../enyu_secs.json';

interface LoginParams {
    username?: string;
    password?: string;
}

interface LoginResponse {
    headers: {
        'set-cookie': string[];
    };
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
    cookies: string[];
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
    if (!username || !password) {
        const auth = secs.auth;
        username = auth.username;
        password = auth.password;
    }
    const auth = {
        is_ajax1: true,
        time_on_page1: 5,
        assets_version1: 1866,
        username,
        password,
        remember: 1,
        f_ajax_login: 1,
    };

    const formData = gs.util.getFormData(auth);
    const res = await gs.util.doHttpRequest({
        method: 'POST',
        url: 'https://freedcamp.com/login',
        data: formData || '',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
    });

    return res as LoginResponse;
}

async function getProcessor(): Promise<Processor> {
    const loginRes = await login({});
    const cookies = loginRes.headers['set-cookie'];

    async function doPostMultiPart(path: string, json: string | object): Promise<any> {
        if (typeof json !== 'string') {
            json = JSON.stringify(json);
        }
        const data = `------geckoformboundarya5436b018dcf600688cc0244d5319984\r\nContent-Disposition: form-data; name="data"\r\n
${json}${
            '\r\n'
}------geckoformboundarya5436b018dcf600688cc0244d5319984--\r\n`;
        const res = await gs.util.doHttpRequest({
            method: 'POST',
            url: `https://freedcamp.com${path}`,
            data,
            headers: {
                Cookie: cookies.join('; '),                
                'Content-Type': 'multipart/form-data; boundary=----geckoformboundarya5436b018dcf600688cc0244d5319984',
            },
        });
        return res;
    }

    async function doPostAttachment(taskId: string, params: TaskParams): Promise<any> {
        await doPostMultiPart(`/iapi/tasks/${taskId}`, { 
            ...params,
            "conditions": { "filter": {}, "order": {}, "substring": "", "f_use_and": "0" }, 
            "time_on_page": 37378 
        });
    }

    async function createTask(projectAndGroup: ProjectAndGroup, title: string): Promise<CreateTaskResponse> {
        const res = await doPostMultiPart('/iapi/tasks',
            {
                "title": title,
                "assigned_to_id": "0", 
                ...projectAndGroup, 
                "conditions": {"filter": {}, "order": {}, "substring": "", "f_use_and": "0"},
                "time_on_page": 3704
            }
        );
        const result = res.data;
        return {
            result,
            id: result.data.tasks[0].id,
        };
    }

    async function doDelAction(path: string): Promise<any> {
        const res = await gs.util.doHttpRequest({
            method: 'DELETE',
            url: `https://freedcamp.com${path}`,
            headers: {
                Cookie: cookies.join('; '),
            },
        });
        return res;
    }

    async function doGetAction(path: string): Promise<any> {
        const res = await gs.util.doHttpRequest({
            method: 'GET',
            url: `https://freedcamp.com${path}`,
            headers: {
                Cookie: cookies.join('; '),
            },
        });
        return res;
    };

    async function deleteTask(taskId: string | number): Promise<any> {
        const res = await doDelAction(`/iapi/tasks/${taskId}`);
        return res;
    }

    async function getSessionCurrentData(): Promise<ICurrentSessionData> {
        const res = await doGetAction('/iapi/sessions/current');
        return res.data as ICurrentSessionData;
    }

    return {
        cookies,
        doPostAttachment,
        createTask,
        deleteTask,
        getSessionCurrentData,
    };
}

export {
    login,
    getProcessor,
    secs,
    LoginParams,
    LoginResponse,
    ProjectAndGroup,
    CreateTaskResponse,
    Processor,
};
