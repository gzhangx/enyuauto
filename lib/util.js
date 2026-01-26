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
            contentType: 'application/x-www-form-urlencoded; charset=UTF-8',
        },
    });

    return res;
}



module.exports = {
    login,
};