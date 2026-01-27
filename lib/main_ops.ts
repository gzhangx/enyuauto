import * as login from './util';
import * as gs from '@gzhangx/googleapi';
import * as fs from 'fs';
import { ProjectAndGroup, Processor, IUserInfo } from './util';

interface Operation {
    '文件': string;
    '作者': string;
    '文章名': string;
    '文章链接': string;
    '作者电邮': string;
    '校对': string;
    '美编': string;
    '发布': string;
    'send': string;
}

interface OperationInfo {
    author: string;
    article: string;
    link: string;
    email: string;
    '校对': string;
    '美编': string;
    '发布': string;
}

interface Templates {
    [key: string]: string;
}

interface OperationAndTemplates {
    validOperations: Operation[];
    templates: Templates;
    ops: gs.gsAccount.IGetSheetOpsReturn;
}

async function getSheetOps(): Promise<gs.gsAccount.IGetSheetOpsReturn> {
    const gsc = await gs.google.gsAccount.getClient(login.secs.gsAuth);
    const ops = await gsc.getSheetOps(login.secs.gsAuth.main_sheet_id);
    return ops;
}

async function getOperationAndTemplates(): Promise<OperationAndTemplates> {    
    const ops = await getSheetOps();
    const operationList = await ops.readDataByColumnName('main');
    const validOperations = (operationList.data || []).filter((d: any) => d['send'] === 'Y') as unknown as Operation[];

    const templateRows = await ops.readData('templates');
    const templates = templateRows.values.reduce((acc: Templates, row: string[]) => {
        const [templateName, ...templateContent] = row;
        if (templateName) {
            acc[templateName] = templateContent.join('\n');
        }
        return acc;
    }, {});
    return { validOperations, templates, ops };
}

function getProjectGroupMapping(): { [key: string]: ProjectAndGroup } {
    return {
        '校对': { "project_id": "3696514", "task_group_id": "6825082" },
        '美编': { "project_id": "3696516", "task_group_id": "6825087" },
        '发布': { "project_id": "3696243", "task_group_id": "6824401" },
    };
}

function getActions(): ('校对' | '美编' | '发布')[] {
    return ['校对', '美编', '发布'];
}

async function processOperation(
    validOperations: Operation[], 
    templates: Templates, 
    debug_Prefix: string = ''
): Promise<number[]> {
    const pr = await login.getProcessor();
    
    const taskIds = [];
    const userData = await pr.getSessionCurrentData();
    const userNameToInfoMap = userData.data.users.reduce((acc, user) => {
        acc[user.full_name] = user;
        return acc;
    }, {} as { [key: string]: IUserInfo });
    for (const operation of validOperations) {
        const fileName = operation['文件'];
        const infos: OperationInfo = {
            author: operation['作者'],
            article: operation['文章名'],
            link: operation['文章链接'],
            email: operation['作者电邮'],
            //校对: operation['校对'],
        } as OperationInfo;

        const projectGroupMapping = getProjectGroupMapping();
        for (const action of getActions()) {
            infos[action] = operation[action];
            let template1 = templates[action];
            const projectGrpoup = projectGroupMapping[action];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id;
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            taskIds.push(taskId);
            const link = infos.link;
            const editor = infos[action];
            for (const replaceItem of [...getActions(), 'author', 'email', 'article']) {
                if (replaceItem === 'article' && link) continue;
                template1 = template1.replace(`{${replaceItem}}`, infos[replaceItem as keyof OperationInfo]);
            }
            if (link) {
                template1 = template1.replace('{article}', `<a href="${infos['link']}">${infos['article']}</a>`);
            }
            let assigned_to_id = '';

            if (editor) {
                const userInfo = userNameToInfoMap[editor];
                if (userInfo) {
                    assigned_to_id = userInfo.user_id;
                }
            }
            //due_date
            await pr.doPostAttachment(taskId, assigned_to_id, template1);
        }
    }
    return taskIds;
}

async function debug_main(opStr?: string): Promise<void> {
    if (opStr === 'del') {
        const ops = await getSheetOps();
        const temp = await ops.readData('temp');
        
        console.log('Deleting tasks:', temp);
        const pr = await login.getProcessor();        
        
        for (const taskId of temp.values[0][0].split(',')) {
            if (taskId) {
                await pr.deleteTask(taskId);
                console.log(`Deleted task ${taskId}`);
            }
        }
        return;
    }
    const { ids, ops } = await main(opStr);
    console.log('Created task IDs:', ids);
    
    await ops.updateValues('temp!A1', [['forceStringValDEBUGNO_USE,'+ids.join(',')]]);
    
    console.log(`Saved ${ids.length} task IDs to temp/taskId.txt`);
}

async function main(opStr?: string) {
    const { validOperations, templates, ops } = await getOperationAndTemplates();
    const prefix = opStr || '';
    const ids = await processOperation(validOperations, templates, prefix);
    return {
        ids, 
        ops,
    }
}

export {
    debug_main,
    main,
};
