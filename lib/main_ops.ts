import * as login from './util';
import * as gs from '@gzhangx/googleapi';
import * as fs from 'fs';
import { ProjectAndGroup, Processor, IUserInfo, TaskParams } from './util';

const actions = ['校对', '美编', '发布'] as const;
type ActionType = typeof actions[number];
type DueDateKeys = `${ActionType} Due Date`;

interface Operation {
    '文件': string;
    '作者': string;
    '文章名': string;
    '文章链接': string;
    '作者电邮': string;
    '文章类别': string;
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
    category: string;
    editor: string;
}


type Templates = {
    [K in ActionType]: {
        template: string;
        taskIdPos: number;
        taskIdLine: number;
        taskIdUpdater: (newTaskId: string) => Promise<void>;
        existingTaskId: string;
    }
};

interface OperationAndTemplates {
    validOperation: OperationWithDueDates;
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

async function getOperationAndTemplates(lineNumber: number, log: DebugLog): Promise<OperationAndTemplates|undefined> {    
    const ops = await getSheetOps();
    log.operations.push('getOperationAndTemplates: got sheet ops');
    const operationList = await ops.readDataByColumnName('main');
    log.operations.push('getOperationAndTemplates: got operation list from main');
    const validOperation = (operationList.data || [])[lineNumber - 2] as unknown as OperationWithDueDates;

    if (!validOperation) {
        log.operations.push('getOperationAndTemplates: no such line number ' + lineNumber);
        return undefined;
    }
    
    log.operations.push('getOperationAndTemplates: got operation ' + validOperation['文件']);
    const templateRows = await ops.readData('templates');
    log.operations.push('getOperationAndTemplates: got template rows');
    const templates = templateRows.values.reduce((acc: Templates, row: string[]) => {
        const [templateName, ...templateContent] = row;
        if (templateName) {
            acc[templateName as ActionType] = {
                template: templateContent.join('\n'),
                taskIdPos: -1, //headers.values[0].indexOf(templateName),
                taskIdLine: lineNumber,
                taskIdUpdater: async (newTaskId: string) => { },
                existingTaskId: validOperation[`${templateName} TaskId` as DueDateKeys]
            };
        }
        return acc;
    }, {} as Templates);
    const headers = await ops.readData('main', { row: 1, col: 20 });
    for (let index = 0; index < headers.values[0].length; index++) {
        const item = headers.values[0][index];
        for (const key of actions) {
            if (item === `${key} TaskId`) {
                log.operations.push(`getOperationAndTemplates: header found ${item} at index ${index}`);
                templates[key].taskIdPos = index;
                templates[key].taskIdUpdater = async (newTaskId: string) => {
                    await ops.autoUpdateValues('main', [[newTaskId]], {
                        row: lineNumber - 1,
                        col: index,
                    });
                }
                //await templates[key].taskIdUpdater('test' + key + "test");
                break;
            }
        }        
    }

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
    return { validOperation, templates, ops, editorInfoMap };
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
): Promise<string[]> {
    const { validOperation, templates, editorInfoMap } = opsAndTemplates;
    const pr = await login.getProcessor();
    log.operations.push('processOperation: got processor with login');
    const taskIds = [];
    const userData = await pr.getSessionCurrentData();
    const userNameToInfoMap = userData.data.users.reduce((acc, user) => {
        acc[user.full_name] = user;
        return acc;
    }, {} as { [key: string]: IUserInfo });
    const operation = validOperation;
    {
        const fileName = operation['文件'];
        const infos: OperationInfo = {
            author: operation['作者'],
            article: operation['文章名'],
            link: operation['文章链接'],
            email: operation['作者电邮'],
            category: operation['文章类别'],
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
            const curTemplateActionInfo = templates[action];
            if (curTemplateActionInfo.existingTaskId) {
                log.operations.push(`processOperation: skipping action ${action} for file ${fileName} as task ID ${curTemplateActionInfo.existingTaskId} exists`);
                taskIds.push(curTemplateActionInfo.existingTaskId);
                continue;
            }
            let template1 = curTemplateActionInfo.template;
            const projectGrpoup = projectGroupMapping[action];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id.toString();
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            log.operations.push(`processOperation: created task action ${action} ${taskId} for file ${fileName}`);
            taskIds.push(taskId);
            const link = action === '校对' ? infos.link : undefined;            
            
            for (const replaceItem of ['editor', 'author', 'email', 'article', 'category']) {
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
async function debug_main(lineNumber: number, log:DebugLog, opStr?: string): Promise<boolean> {
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
        return true;
    }
    const mainResult = await main(lineNumber, log, opStr);
    if (!mainResult) {        
        return false;
    }
    const { ids, ops } = mainResult;
    log.operations.push('Created task IDs:'+ JSON.stringify(ids));
    
    await ops.updateValues('temp!A1', [['forceStringValDEBUGNO_USE,'+ids.join(',')]]);
    
    log.operations.push(`Saved ${ids.length} task IDs to temp/taskId.txt`);
    return true;
}

async function main(lineNumber: number, log: DebugLog, opStr?: string) {
    const opsAndTemplates = await getOperationAndTemplates(lineNumber, log);
    if (!opsAndTemplates) {        
        return undefined;
    }
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
