const login = require('./lib/util');
const https = require('https');
const secs = require('./secs.json');
const gs = require('@gzhangx/googleapi');

async function test() {
    const pr = await login.getProcessor();
    await testResponse(pr.cookies.map(c => c).join('; '), 'from https request');
    //await pr.doPostAttachment('/iapi/tasks/69187618', 'this is a test attachment from test').then(res => {
    //    console.log('Attachment upload response status(from gs util):', res.statusCode, res.data);
    //});
    //return;
    //const res = await login.login({});
    //console.log(res.data.toString());
    //console.log(res.statusMessage);
    //console.log(res.headers)
    //console.log('Test done', pr.cookies);




    
    async function testResponse(cookie, desc) {      
const data = `------geckoformboundarya5436b018dcf600688cc0244d5319984\r\nContent-Disposition: form-data; name="data"\r\n
{"description":"${desc}","conditions":{"filter":{},"order":{},"substring":"","f_use_and":"0"},"time_on_page":37378}${
            '\r\n'
}------geckoformboundarya5436b018dcf600688cc0244d5319984--\r\n`
      const res = await gs.util.doHttpRequest({
                  method: 'POST',
          url: `https://freedcamp.com/iapi/tasks/69187618`,
                  data,
                  headers: {
                      //Cookie: cookies.join('; '),
                      //cookie: 'AWSALB=436tpJ0fCGT270ai4gZxektI8bF8rK7IaIdK55KpxgPpAiD+zJnh430vYafOBuvNK1qMmS0nWZEDcJDAFXj04dVLBAkqx1ux+KpY9qsj9UIh1zdNwSw87pg0TS/P; AWSALBCORS=436tpJ0fCGT270ai4gZxektI8bF8rK7IaIdK55KpxgPpAiD+zJnh430vYafOBuvNK1qMmS0nWZEDcJDAFXj04dVLBAkqx1ux+KpY9qsj9UIh1zdNwSw87pg0TS/P; fc_lang=en; fcrel_new_frontpage_c=0; fcrel_new_frontpage=1; _ga_DXYN2SN4L9=GS2.1.s1769443065$o2$g1$t1769443352$j60$l0$h0; _ga=GA1.1.369532003.1769404261; fc_vz=4330521f4738c8cd37e7798602213730_8701927; remember_identifier=04eaf8aa68ee9c9525e5bfce788eaaf25b528ea7; remember_code=accbab597e16bff5e03e1f223f406060a06e9df1; identity=user_1320079; ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                      //cookie: 'remember_identifier=04eaf8aa68ee9c9525e5bfce788eaaf25b528ea7; remember_code=accbab597e16bff5e03e1f223f406060a06e9df1; identity=user_1320079; ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                      //cookie: ' ci_session=6S6UuGGqnJDjglB%2C2xSLiKAFoySMxQY7U6tOdOigm-OzYRUR',
                      //cookie: 'ci_session=K6V18H%2Cq5981kxCGXoh48Em5sQjGZVkCTqQgJtGCFvwf%2CUYb',
                      //cookie-gd: 'ci_session=K6V18H%2Cq5981kxCGXoh48Em5sQjGZVkCTqQgJtGCFvwf%2CUYb',
                      //cookie-bd: 'ci_session=oLXKkKXPlb3lmAaJhfNE5F6kFf5lxLipClW3I8EwlN6vKs7e',
                      //cookie-bd: 'ci_session=k9KDXcRKzj7S0ulwSSu39NSV4l76L2rIo8oiKE1ReC6p-T0R'
                      cookie,
                      'Content-Type': 'multipart/form-data; boundary=----geckoformboundarya5436b018dcf600688cc0244d5319984',
                  },
      
      });
      console.log(`Attachment upload response status(dec=${desc}):`, res.statusCode, res.data);
}

    const params = gs.util.getFormData({
        is_ajax1: true,
        time_on_page1: 5,
        assets_version1: 1866,
        username: secs.auth.username,
        password: secs.auth.password,
        remember: 1,
        //a_token1:6976fe604ecbb,
        f_ajax_login: 1,
    });
const postData = params;

const options = {
  hostname: 'freedcamp.com',
  path: '/login',
  method: 'POST',
  headers: {
    //'Cookie': 'AWSALB=j3lD80hxEfAwdF1fgaT7Mfdre7kPDSDlvydgKhMztJd97u0qA0OJHP5cmSq5ctWnzEnsarMjVDpkCGKgueuzx4nKTKqcx1hfqLkbJNXmsC0dnIXCWvrs4tGjLqCl; AWSALBCORS=j3lD80hxEfAwdF1fgaT7Mfdre7kPDSDlvydgKhMztJd97u0qA0OJHP5cmSq5ctWnzEnsarMjVDpkCGKgueuzx4nKTKqcx1hfqLkbJNXmsC0dnIXCWvrs4tGjLqCl; ci_session=w9UfRAXCm%2CNYKHzI9NL2B0DKdS42QH5OJS-rnqCWI3gBsVrk; fc_lang=en; identity=user_1320079; remember_code=533e89772736c3a5ca40ae5c1ea89ed3c3b66250; remember_identifier=7b32068f59280b2caab787b0b7d3be40c6d05de8',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Content-Length': Buffer.byteLength(postData)
  }
};

const req = https.request(options, (response) => {
  const ck = response.headers['set-cookie'];
  console.log('ck from https login:', ck);
  
  let body = '';
  response.on('data', (chunk) => {
    body += chunk;
  });
  
  response.on('end', async () => {
      testResponse(ck.map(c => c).join('; '), 'from https request');
  });
});

req.on('error', (error) => {
  console.log(error);
});

req.write(postData);
req.end();

}

test();