const login = require('./lib/util');
const gs = require('@gzhangx/googleapi');

const fs = require('fs');
async function getOperationAndTemplates() {
    const gsc = await gs.google.gsAccount.getClient(login.secs.gsAuth);
    const ops = await gsc.getSheetOps(login.secs.gsAuth.main_sheet_id);
    const operationList = await ops.readDataByColumnName('main');
    const validOperations = operationList.data.filter(d => d['send'] === 'Y');

    const templateRows = await ops.readData('templates');
    const templates = templateRows.values.reduce((acc, row) => {
        const [templateName, ...templateContent] = row;
        if (templateName) {
            acc[templateName] = templateContent.join('\n');
        }
        return acc;
    }, {});
    return { validOperations, templates };
}


function getProjectGroupMapping() {
    return {
        '校对': { "project_id": "3696514", "task_group_id": "6825082" },
        '美编': { "project_id": "3696516", "task_group_id": "6825087" },
        '发布': { "project_id": "3696243", "task_group_id": "6824401" },
    }
}

async function processOperation(validOperations, templates, debug_Prefix = '') {
    const pr = await login.getProcessor();
    console.log('Valid operations:', validOperations, templates);
    for (const operation of validOperations) {
        const fileName = operation['文件'];
        const infos = {
            author: operation['作者'],
            article: operation['文章名'],
            link: operation['文章链接'],
            email: operation['作者电邮'],
            editor: operation['校对'],
        }

        const projectGroupMapping = getProjectGroupMapping();
        for (const actions of ['校对', '美编', '发布']) {
            let template1 = templates['校对'];
            const projectGrpoup = projectGroupMapping[actions];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id;
            console.log(`Created task ${taskId} for file ${fileName}`);
            fs.writeFileSync('./temp/taskId.txt', taskId.toString());
            const link = infos.link;
            for (const replaceItem of ['editor', 'author', 'email', 'article']) {
                if (replaceItem === 'article' && link) continue;
                template1 = template1.replace(`{${replaceItem}}`, infos[replaceItem]);
            }
            if (link) {
                template1 = template1.replace('{article}', `<a href="${infos['link']}">${infos['article']}</a>`);
            }
            await pr.doPostAttachment(taskId, template1);
        }
    }
}
    


async function main() {
    if (process.argv.includes('del')) {
        const pr = await login.getProcessor();
        const taskId = fs.readFileSync('./temp/taskId.txt', 'utf8').trim();
        await pr.deleteTask(taskId);
        console.log(`Deleted task ${taskId}`);
        return;
    }
    const { validOperations, templates } = await getOperationAndTemplates();
    await processOperation(validOperations, templates, 'test_');
}

main().catch(err => {
    console.error('Test failed:', err);
});