const login = require('./lib/util');
const gs = require('@gzhangx/googleapi');

const fs = require('fs');
async function test() {


    const gsc = await gs.google.gsAccount.getClient(login.secs.gsAuth);
    const ops = await gsc.getSheetOps(login.secs.gsAuth.main_sheet_id);
    const d = await ops.readDataByColumnName('main');
    return console.log(d);
    const pr = await login.getProcessor();
    
    const args = process.argv.slice(2);
    if (args.includes('del')) {
        const taskId = fs.readFileSync('./temp/taskId.txt', 'utf8').trim();
        await pr.deleteTask(taskId);
        console.log(`Deleted task ${taskId}`);
        return;
    }
    
    const taskRes = await pr.createTask('testtask');
    const taskId = taskRes.id;
    fs.writeFileSync('./temp/taskId.txt', taskId.toString());
    //await testResponse(pr.cookies.map(c => c).join('; '), 'from https request');
    await pr.doPostAttachment(taskId, 'This is a test attachment upload.');
    //return;
    //const res = await login.login({});
    //console.log(res.data.toString());
    //console.log(res.statusMessage);
    //console.log(res.headers)
    //console.log('Test done', pr.cookies);




    
    async function testResponse(cookie, desc) {
        const data = `------geckoformboundarya5436b018dcf600688cc0244d5319984\r\nContent-Disposition: form-data; name="data"\r\n
{"description":"${desc}","conditions":{"filter":{},"order":{},"substring":"","f_use_and":"0"},"time_on_page":37378}${'\r\n'
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

}



test().catch(err => {
    console.error('Test failed:', err);
});