import * as gs from '@gzhangx/googleapi';

import type { FreedCampLoginParams, LoginResponse, ProjectTaskParams, CreateTaskResponse, FreedCampProcessor, ICurrentSessionData, FreedCampOps, IProjectTasksResult} from '../web/shared/freedcampTypes';


async function login({ username, password }: FreedCampLoginParams): Promise<LoginResponse> {
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

    const rsp = res as {
        headers: {
            'set-cookie': string[];
        };
    };
    return { Cookie: rsp.headers['set-cookie'].join('; ') };
}

function getProcessor(loginResponse: LoginResponse): FreedCampProcessor {
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
        console.log('DEBUGsubmit data', {
            method: 'POST',
            url: `https://freedcamp.com${path}`,
            data,
            headers: {
                Cookie: loginResponse.Cookie,                
                'Content-Type': 'multipart/form-data; boundary=----geckoformboundarya5436b018dcf600688cc0244d5319984',
            },
        })
        const res = await gs.util.doHttpRequest({
            method: 'POST',
            url: `https://freedcamp.com${path}`,
            data,
            headers: {
                Cookie: loginResponse.Cookie,                
                'Content-Type': 'multipart/form-data; boundary=----geckoformboundarya5436b018dcf600688cc0244d5319984',
            },
        });
        console.log('DEBUGsubmit resdata ', res.data?.http_code,res.data?.msg,res.data?.data)
        return res;
    }

    async function doPostAttachment(taskId: string, params: ProjectTaskParams): Promise<any> {
        return await doPostMultiPart(`/iapi/tasks/${taskId}`, { 
            ...params,
            "conditions": { "filter": {}, "order": {}, "substring": "", "f_use_and": "0" }, 
            "time_on_page": 37378 
        });
    }

    async function createTask(projectAndGroup: ProjectTaskParams, title: string): Promise<CreateTaskResponse> {
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
        console.log('DEBUGsubmit create task resdata ', res.data?.http_code,res.data?.msg,res.data?.data)
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
                Cookie: loginResponse.Cookie,
            },
        });
        return res.data;
    }

    async function doGetAction(path: string): Promise<any> {
        const res = await gs.util.doHttpRequest({
            method: 'GET',
            url: `https://freedcamp.com${path}`,
            headers: {
                Cookie: loginResponse.Cookie,
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

    async function getTasksForProjects(projectId: string, pageNumber: number = 1): Promise<IProjectTasksResult> {
        const res = await doGetAction(`/iapi/tasks?project_id=${
            projectId
        }&page_num=${pageNumber}&filter={}&order={}&substring=""&f_use_and=0&f_react_app=1&f_include_tr_data=1&f_include_tags=0&f_include_ms_data=true&group_mode=lists&group_mode_tpl_id=`);
        return res as IProjectTasksResult;
    }

    return {
        doPostAttachment,
        createTask,
        deleteTask,
        getSessionCurrentData,
        getTasksForProjects,
    };
}

export const freedCampOps: FreedCampOps = {
    login,
    getProcessor,
}
