const login = require('./util');
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
function getActions() {
    return ['校对', '美编', '发布'];
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
        for (const action of getActions()) {
            let template1 = templates[action];
            const projectGrpoup = projectGroupMapping[action];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id;
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            fs.writeFileSync(`./temp/taskId_${action}.txt`, taskId.toString());
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
        for (const action of getActions()) {
            const taskId = fs.readFileSync(`./temp/taskId_${action}.txt`, 'utf8').trim();
            await pr.deleteTask(taskId);
            console.log(`Deleted task ${taskId} for action ${action}`);
        }
        return;
    }
    const { validOperations, templates } = await getOperationAndTemplates();
    const prefix = process.argv.includes('test') ? 'test_' : '';
    await processOperation(validOperations, templates, prefix);
}

module.exports = {
    main,
}