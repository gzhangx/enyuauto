import * as login from './util';
import * as gs from '@gzhangx/googleapi';
import * as fs from 'fs';
import { ProjectAndGroup, Processor, IUserInfo, TaskParams } from './util';

type ActionType = '校对' | '美编' | '发布';
type DueDateKeys = `${ActionType} Due Date`;

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

type OperationWithDueDates = Operation & {
    [K in DueDateKeys]: string;
};

interface OperationInfo {
    author: string;
    article: string;
    link: string;
    email: string;
    editor: string;
}

type Templates = {
    [K in ActionType]: string;
};

interface OperationAndTemplates {
    validOperations: OperationWithDueDates[];
    templates: Templates;
    ops: gs.gsAccount.IGetSheetOpsReturn;
    editorInfoMap: { [key: string]: IEditorInfo };
}


interface IEditorInfo {
    title: string;
    full_name: string;
    email: string;
    task: string;
    print_name: string; //the full chinese name for printing
}
async function getSheetOps(): Promise<gs.gsAccount.IGetSheetOpsReturn> {
    const gsc = await gs.google.gsAccount.getClient(login.secs.gsAuth);
    const ops = await gsc.getSheetOps(login.secs.gsAuth.main_sheet_id);
    return ops;
}

async function getOperationAndTemplates(log: DebugLog): Promise<OperationAndTemplates> {    
    const ops = await getSheetOps();
    log.operations.push('getOperationAndTemplates: got sheet ops');
    const operationList = await ops.readDataByColumnName('main');
    log.operations.push('getOperationAndTemplates: got operation list from main');
    const validOperations = (operationList.data || []).filter((d: any) => d['send'] === 'Y') as unknown as OperationWithDueDates[];

    log.operations.push('getOperationAndTemplates: got operation list ' + validOperations.length);
    const templateRows = await ops.readData('templates');
    log.operations.push('getOperationAndTemplates: got template rows');
    const templates = templateRows.values.reduce((acc: any, row: string[]) => {
        const [templateName, ...templateContent] = row;
        if (templateName) {
            acc[templateName] = templateContent.join('\n');
        }
        return acc;
    }, {}) as Templates;

    const listOfNames = await ops.readDataByColumnName('list of names');
    log.operations.push('getOperationAndTemplates: got list of names');
    const editorInfoMap = listOfNames.data?.reduce((acc, item) => {
        const full_name = item['Name on FreedCamp'];
        const email = item['Email'];
        const task = item['Task'];
        const title = item['Title'];
        const print_name = item['Full Name'];
        acc[full_name] = { title, full_name, email, task, print_name };
        return acc;
    }, {} as { [key: string]: IEditorInfo }) || {};
    return { validOperations, templates, ops, editorInfoMap };
}

function getProjectGroupMapping(): { [key: string]: ProjectAndGroup } {
    return {
        '校对': { "project_id": "3696514", "task_group_id": "6825082" },
        '美编': { "project_id": "3696516", "task_group_id": "6825087" },
        '发布': { "project_id": "3696243", "task_group_id": "6824401" },
    };
}

function getActions(): ActionType[] {
    return ['校对', '美编', '发布'];
}

async function processOperation(
    opsAndTemplates: OperationAndTemplates,        
    log: DebugLog,
    debug_Prefix: string = ''
): Promise<number[]> {
    const { validOperations, templates, editorInfoMap } = opsAndTemplates;
    const pr = await login.getProcessor();
    log.operations.push('processOperation: got processor with login');
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
            editor: '',
        } as OperationInfo;

        const projectGroupMapping = getProjectGroupMapping();
        for (const action of getActions()) {
            const editor = operation[action];
            const editorLookup = editorInfoMap[editor];
            if (editorLookup) {
                const prettyName = editorLookup.print_name || editorLookup.full_name; 
                if (editorLookup.title === 'Brother') {
                    infos['editor'] = `${editorLookup.title} ${prettyName}`;
                } else {
                    infos['editor'] = `${prettyName}${editorLookup.title}`;
                }
            }
            let template1 = templates[action];
            const projectGrpoup = projectGroupMapping[action];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id;
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            log.operations.push(`processOperation: created task action ${action} ${taskId} for file ${fileName}`);
            taskIds.push(taskId);
            const link = action === '校对' ? infos.link : undefined;            
            
            for (const replaceItem of ['editor', 'author', 'email', 'article']) {
                if (replaceItem === 'article' && link) continue;
                template1 = template1.replace(`{${replaceItem}}`, infos[replaceItem as keyof OperationInfo]);
            }
            if (link) {
                template1 = template1.replace('{article}', `<a href="${infos['link']}">${infos['article']}</a>`);
            }
            
            const postParams: TaskParams = {
                description: template1,
            };
            const dueDateKey = `${action} Due Date` as DueDateKeys;
            const due_date = operation[dueDateKey];

            if (due_date) {
                postParams.due_date = due_date;
            }
            if (editor) {
                const userInfo = userNameToInfoMap[editor];
                if (userInfo) {
                    //DEBUG don't assign it sends email
                    //postParams.assigned_to_id = userInfo.user_id;
                    if (debug_Prefix) {
                        //postParams.assigned_to_id = '1320079';
                    }
                }
            }
            log.operations.push(`processOperation: prepared post params for task ${taskId} action ${action} with ${JSON.stringify(postParams)}`);
            //due_date
            await pr.doPostAttachment(taskId, postParams);
        }
    }
    return taskIds;
}

export interface DebugLog {
    operations: string[];
}
async function debug_main(log:DebugLog, opStr?: string): Promise<void> {
    if (opStr === 'del') {
        const ops = await getSheetOps();
        log.operations.push('del: got ops');
        const temp = await ops.readData('temp');
        log.operations.push('del: got temp data');
        
        //log.operations.push('Deleting tasks:', temp);        
        const pr = await login.getProcessor();        
        
        for (const taskId of temp.values[0][0].split(',')) {
            if (taskId) {
                log.operations.push(`del: deleting task ${taskId}`);
                await pr.deleteTask(taskId);
                log.operations.push(`del: deleted task ${taskId}`);
            }
        }
        return;
    }
    const { ids, ops } = await main(log,opStr);
    log.operations.push('Created task IDs:'+ JSON.stringify(ids));
    
    await ops.updateValues('temp!A1', [['forceStringValDEBUGNO_USE,'+ids.join(',')]]);
    
    log.operations.push(`Saved ${ids.length} task IDs to temp/taskId.txt`);
}

async function main(log: DebugLog, opStr?: string) {
    const opsAndTemplates = await getOperationAndTemplates(log);
    const prefix = opStr || '';
    const ids = await processOperation(opsAndTemplates, log, prefix);
    return {
        ids, 
        ...opsAndTemplates,
    }
}

export {
    debug_main,
    main,
};
