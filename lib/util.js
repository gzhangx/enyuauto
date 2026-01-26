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

    async function doPostAttachment(path, text) {
        const data = `------geckoformboundarya5436b018dcf600688cc0244d5319984\r\nContent-Disposition: form-data; name="data"\r\n
{"description":"${text}","conditions":{"filter":{},"order":{},"substring":"","f_use_and":"0"},"time_on_page":37378}${
            '\r\n'
}------geckoformboundarya5436b018dcf600688cc0244d5319984--\r\n`
        const res = await gs.util.doHttpRequest({
            method: 'POST',
            url: `https://freedcamp.com${path}`,
            data,
            headers: {
                Cookie: cookies.join('; '),
                //cookie: 'AWSALB=436tpJ0fCGT270ai4gZxektI8bF8rK7IaIdK55KpxgPpAiD+zJnh430vYafOBuvNK1qMmS0nWZEDcJDAFXj04dVLBAkqx1ux+KpY9qsj9UIh1zdNwSw87pg0TS/P; AWSALBCORS=436tpJ0fCGT270ai4gZxektI8bF8rK7IaIdK55KpxgPpAiD+zJnh430vYafOBuvNK1qMmS0nWZEDcJDAFXj04dVLBAkqx1ux+KpY9qsj9UIh1zdNwSw87pg0TS/P; fc_lang=en; fcrel_new_frontpage_c=0; fcrel_new_frontpage=1; _ga_DXYN2SN4L9=GS2.1.s1769443065$o2$g1$t1769443352$j60$l0$h0; _ga=GA1.1.369532003.1769404261; fc_vz=4330521f4738c8cd37e7798602213730_8701927; remember_identifier=04eaf8aa68ee9c9525e5bfce788eaaf25b528ea7; remember_code=accbab597e16bff5e03e1f223f406060a06e9df1; identity=user_1320079; ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                //cookie: 'remember_identifier=04eaf8aa68ee9c9525e5bfce788eaaf25b528ea7; remember_code=accbab597e16bff5e03e1f223f406060a06e9df1; identity=user_1320079; ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                //cookie: ' ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                //cookie: 'ci_session=K6V18H%2Cq5981kxCGXoh48Em5sQjGZVkCTqQgJtGCFvwf%2CUYb',
                //cookie-gd: 'ci_session=K6V18H%2Cq5981kxCGXoh48Em5sQjGZVkCTqQgJtGCFvwf%2CUYb',
                //cookie-bd: 'ci_session=oLXKkKXPlb3lmAaJhfNE5F6kFf5lxLipClW3I8EwlN6vKs7e',
                //cookie-bd: 'ci_session=k9KDXcRKzj7S0ulwSSu39NSV4l76L2rIo8oiKE1ReC6p-T0R'
                //   cookie: 'ci_session=w9UfRAXCm%2CNYKHzI9NL2B0DKdS42QH5OJS-rnqCWI3gBsVrk',
                //           "ci_session=k3bSBpSeirolme-cb2reC7Y-6IP4bnijatA%2CL2ecGhB4zhhk; expires=Mon, 26-Jan-2026 22:11:32 GMT; Max-Age=7200; path=/; samesite=Lax; secure; HttpOnly"
                'Content-Type': 'multipart/form-data; boundary=----geckoformboundarya5436b018dcf600688cc0244d5319984',
            },

        });
        return res;
    }
    return {
        cookies,
        doPostAttachment
    };
}



module.exports = {
    login,
    getProcessor,
};