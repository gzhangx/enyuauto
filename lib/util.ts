import * as gs from '@gzhangx/googleapi';

import type {LoginParams, LoginResponse, ActionTaskParams, CreateTaskResponse, FreedCampProcessor, ICurrentSessionData} from '../web/src/lib/shared/freedcampTypes';


async function login({ username, password }: LoginParams): Promise<string[]> {
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

    const rsp = res as LoginResponse;
    return rsp.headers['set-cookie'];
}

function getProcessor(cookies: string[]): FreedCampProcessor {
    //const loginRes = await login({});
    //const cookies = loginRes.headers['set-cookie'];

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

    async function doPostAttachment(taskId: string, params: ActionTaskParams): Promise<any> {
        await doPostMultiPart(`/iapi/tasks/${taskId}`, { 
            ...params,
            "conditions": { "filter": {}, "order": {}, "substring": "", "f_use_and": "0" }, 
            "time_on_page": 37378 
        });
    }

    async function createTask(projectAndGroup: ActionTaskParams, title: string): Promise<CreateTaskResponse> {
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
        return res.data;
    }

    async function doGetAction(path: string): Promise<any> {
        const res = await gs.util.doHttpRequest({
            method: 'GET',
            url: `https://freedcamp.com${path}`,
            headers: {
                Cookie: cookies.join('; '),
            },
        });
        return res.data;
    };

    async function deleteTask(taskId: string | number): Promise<any> {
        const res = await doDelAction(`/iapi/tasks/${taskId}`);
        return res;
    }

    async function getSessionCurrentData(): Promise<ICurrentSessionData> {
        const res = await doGetAction('/iapi/sessions/current');
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
    LoginParams,
    LoginResponse,
    ActionTaskParams,
    CreateTaskResponse,
};
