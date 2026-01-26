const login = require('./lib/util');

async function test() {
    const res = await login.login({});
    console.log(res.data.toString());
    console.log(res.statusMessage);
    console.log(res.headers)
}

test();