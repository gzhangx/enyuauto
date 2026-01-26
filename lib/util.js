const gs = require('@gzhangx/googleapi');

const secs = require('../secs.json');
async function login({ username, password }) {
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
        //a_token1:6976fe604ecbb,
        f_ajax_login: 1,
    };

    console.log('Logging in with username:', gs.util.getFormData(auth));
    const res = await gs.util.doHttpRequest({
        method: 'POST',
        url: 'https://freedcamp.com/login',
        
        data: gs.util.getFormData(auth),
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
    });

    return res;
}

async function getProcessor() {
    const loginRes = await login({});
    const cookies = loginRes.headers['set-cookie'];

    async function doPostMultiPart(path, json) {
        if (typeof json !== 'string') {
            json = JSON.stringify(json);
        }
        const data = `------geckoformboundarya5436b018dcf600688cc0244d5319984\r\nContent-Disposition: form-data; name="data"\r\n
${json}${
            '\r\n'
}------geckoformboundarya5436b018dcf600688cc0244d5319984--\r\n`
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

    async function doPostAttachment(taskId, description) {
        await doPostMultiPart(`/iapi/tasks/${taskId}`, { description, "conditions": { "filter": {}, "order": {}, "substring": "", "f_use_and": "0" }, "time_on_page": 37378 });
    }
    async function createTask(title) {
        const res = await doPostMultiPart('/iapi/tasks',
            {"title":title,"assigned_to_id":"0","project_id":"3696514","task_group_id":"6825082","conditions":{"filter":{},"order":{},"substring":"","f_use_and":"0"},"time_on_page":3704}
        )
        const result = res.data;
        return {
            result,
            id: result.data.tasks[0].id,
        }
    }

    async function doDelAction(path) {
        const res = await gs.util.doHttpRequest({
            method: 'DELETE',
            url: `https://freedcamp.com${path}`,
            headers: {
                Cookie: cookies.join('; '),
            },

        });
        return res;
    }

    async function deleteTask(taskId) {
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



module.exports = {
    login,
    getProcessor,
};