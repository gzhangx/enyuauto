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
    project_id: string;
    task_group_id: string;
}

interface CreateTaskResponse {
    result: any;
    id: number;
}

interface Processor {
    cookies: string[];
    doPostAttachment: (taskId: number, description: string) => Promise<any>;
    createTask: (projectAndGroup: ProjectAndGroup, title: string) => Promise<CreateTaskResponse>;
    deleteTask: (taskId: string | number) => Promise<any>;
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

    async function doPostAttachment(taskId: number, description: string): Promise<any> {
        await doPostMultiPart(`/iapi/tasks/${taskId}`, { 
            description, 
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

    async function deleteTask(taskId: string | number): Promise<any> {
        const res = await doDelAction(`/iapi/tasks/${taskId}`);
        return res;
    }

    return {
        cookies,
        doPostAttachment,
        createTask,
        deleteTask,
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
