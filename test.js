//const login = require('./lib/util');

async function test() {
    //const pr = await login.getProcessor();
    //await pr.doPostAttachment('/iapi/tasks/69187618', 'this is a test attachment from test').then(res => {
    //    console.log('Attachment upload response status:', res.statusCode, res.data);
    //});
    //const res = await login.login({});
    //console.log(res.data.toString());
    //console.log(res.statusMessage);
    //console.log(res.headers)
    //console.log('Test done', pr.cookies);


const axios = require('axios');
const FormData = require('form-data');
    let data = new FormData();
    const secs = require('./secs.json');
data.append('is_ajax1', 'true');
data.append('time_on_page1', '5');
data.append('assets_version1', '1866');
data.append('username', secs.auth.username);
data.append('password', secs.auth.password);
data.append('remember', '1');
data.append('a_token1', '6976fe604ecbb');
data.append('f_ajax_login', '1');

let config = {
  method: 'post',
  maxBodyLength: Infinity,
  url: 'https://freedcamp.com/login',
  headers: { 
    //'Cookie': 'AWSALB=j3lD80hxEfAwdF1fgaT7Mfdre7kPDSDlvydgKhMztJd97u0qA0OJHP5cmSq5ctWnzEnsarMjVDpkCGKgueuzx4nKTKqcx1hfqLkbJNXmsC0dnIXCWvrs4tGjLqCl; AWSALBCORS=j3lD80hxEfAwdF1fgaT7Mfdre7kPDSDlvydgKhMztJd97u0qA0OJHP5cmSq5ctWnzEnsarMjVDpkCGKgueuzx4nKTKqcx1hfqLkbJNXmsC0dnIXCWvrs4tGjLqCl; ci_session=w9UfRAXCm%2CNYKHzI9NL2B0DKdS42QH5OJS-rnqCWI3gBsVrk; fc_lang=en; identity=user_1320079; remember_code=533e89772736c3a5ca40ae5c1ea89ed3c3b66250; remember_identifier=7b32068f59280b2caab787b0b7d3be40c6d05de8', 
    ...data.getHeaders()
  },
  data : data
};

axios.request(config)
    .then((response) => {
        const ck = response.headers['set-cookie'];
    console.log('ck from axios login:', ck);
  console.log(JSON.stringify(response.data));
})
.catch((error) => {
  console.log(error);
});

}

test();