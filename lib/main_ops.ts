import * as login from './util';
import * as gs from '@gzhangx/googleapi';
import * as fs from 'fs';
import { ProjectAndGroup, Processor, IUserInfo, TaskParams, ICurrentSessionData } from './util';
import { ActionType } from './types';


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
    '二校': string;
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
        templateEnglish: string;
        taskIdPos: number;
        taskIdLine: number;
        taskIdUpdater: (newTaskId: string) => Promise<void>;
        existingTaskId: string;
    }
};

interface IGroupAndMainProjectLongToShortNameMapping {
    groupName: string; //EnYu_2026
    actions: ActionType[];
    taskLongToShortNameMapping: {
        [key: string]: {
            shortName: ActionType; //文字校对 (Editorial and Translation team) : 校对
            subTaskOf?: ActionType;
        }; 
    };
    shortProjectNameToProjectId: { //populated later after we login to freedcamp
        [key in ActionType]: { project_id: string; }; //'校对': { "project_id": "3696514" }
    }
}

interface IEditorInfoMap { [key: string]: IEditorInfo };
interface OperationAndTemplates {
    validOperation: OperationWithDueDates;
    templates: Templates;
    ops: gs.gsAccount.IGetSheetOpsReturn;
    editorInfoMap: IEditorInfoMap;
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
}


interface IEditorInfo {
    title: string;
    shortName: string;
    email: string;
    task: string;
    print_name: string; //the full chinese name for printing
}

interface IOpsConfig {
    operationList: OperationWithDueDates[];
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
    editorInfoMap: IEditorInfoMap;
    headers: string[];
}

async function getSheetOps(): Promise<gs.gsAccount.IGetSheetOpsReturn> {
    const gsc = await gs.google.gsAccount.getClient(login.secs.gsAuth);
    const ops = await gsc.getSheetOps(login.secs.gsAuth.main_sheet_id);
    return ops;
}


export async function getOpsAndMainList(log: DebugLog): Promise<IOpsConfig & { ops: gs.gsAccount.IGetSheetOpsReturn }> {
    const ops = await getSheetOps();
    log.doLog('getOpsAndMainList: got sheet ops');
    const rawMainData = await ops.readData('main');
    const headers = rawMainData.values[0];
    //const operationListData = await ops.readDataByColumnName('main');
    //const operationList = (operationListData.data || []) as unknown as OperationWithDueDates[];
    const operationList: OperationWithDueDates[] = rawMainData.values.slice(1).map(row => {
        const obj = {} as OperationWithDueDates;
        headers.forEach((header, index) => {
            obj[header as keyof OperationWithDueDates] = row[index] || '';
        });
        return obj;
    });
    log.doLog('getOpsAndLine: got operation list from main ' + operationList?.length + ' items');
    const configLines = await ops.readData('config');
    log.doLog('getOperationAndTemplates: got config lines ' + configLines.values.length);
    const groupAndMainProjectMapping = getConfigMapping(configLines.values, log);
    
    const editorInfoMap = getEditorInfoMap(configLines.values);
    return { ops, operationList,groupAndMainProjectMapping,editorInfoMap, headers };
}

async function getOpsAndLine(lineNumber: number, log: DebugLog): Promise<{ ops: gs.gsAccount.IGetSheetOpsReturn, validOperation: OperationWithDueDates, templates: Templates, groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping, editorInfoMap: IEditorInfoMap } | undefined> {
    const { ops, operationList, groupAndMainProjectMapping, editorInfoMap, headers } =  await getOpsAndMainList(log);
    const validOperation = operationList[lineNumber - 2];

    if (!validOperation) {
        log.doLog('getOpsAndLine: no such line number ' + lineNumber);
        return undefined;
    }
    
    log.doLog('getOpsAndLine: got operation ' + validOperation['文件']);
    const templateRows = await ops.readData('templates');
    log.doLog('getOpsAndLine: got template rows');
    const templates = templateRows.values.reduce((acc: Templates, row: string[]) => {
        const [templateName, templateChinese, templateEnglish] = row;
        if (templateName) {
            acc[templateName as ActionType] = {
                template: templateChinese,
                templateEnglish,
                taskIdPos: -1, //headers.values[0].indexOf(templateName),
                taskIdLine: lineNumber,
                taskIdUpdater: async (newTaskId: string) => {
                    const warMsg = 'taskIdUpdater not initialized yet for template ' + templateName;
                    console.warn(warMsg);
                    log.doLog(warMsg);
                 },
                existingTaskId: validOperation[`${templateName} TaskId` as DueDateKeys]
            };
        }
        return acc;
    }, {} as Templates);
    //const headers = await ops.readData('main');
    for (let index = 0; index < headers.length; index++) {
        const item = headers[index];
        for (const key of groupAndMainProjectMapping.actions) {
            if (item === `${key} TaskId`) {
                log.doLog(`getOpsAndLine: header found ${item} at index ${index}`);
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
    
    return { ops, validOperation, templates,groupAndMainProjectMapping, editorInfoMap };
}

async function getOperationAndTemplates(lineNumber: number, log: DebugLog): Promise<OperationAndTemplates|undefined> {    
    const result = await getOpsAndLine(lineNumber, log);
    if (!result) {
        return undefined;
    }
    const { ops, validOperation, templates, editorInfoMap, groupAndMainProjectMapping } = result;

    // const listOfNames = await ops.readDataByColumnName('list of names');
    // log.doLog('getOperationAndTemplates: got list of names');
    // const editorInfoMap = listOfNames.data?.reduce((acc, item) => {
    //     const full_name = item['Name on FreedCamp'];
    //     const email = item['Email'];
    //     const task = item['Task'];
    //     const title = item['Title'];
    //     const print_name = item['Full Name'];
    //     acc[full_name] = { title, full_name, email, task, print_name };
    //     return acc;
    // }, {} as IEditorInfoMap) || {};

    
    return { validOperation, templates, ops, editorInfoMap,groupAndMainProjectMapping };
}


function getEditorInfoMap(values: string[][]): IEditorInfoMap {
    const editorInfoMap: IEditorInfoMap = {};
    const positionToNameMap: { [key: number]: string } = {};
    (values.find(row => row[0] === 'EditorNameColumn') || []).forEach((name, index) => {
        switch (name) {
            case 'Full Name':
                positionToNameMap[index] = 'print_name';
                break;
            case 'Email':
                positionToNameMap[index] = 'email';
                break;
            case 'Task':
                positionToNameMap[index] = 'task';
                break;
            case 'Name on FreedCamp':
                positionToNameMap[index] = 'shortName';
                break;
            case 'Title':
                positionToNameMap[index] = 'title';
                break;
        }
    });
    for (const row of values) {
        if (row[0] === 'EditorName') {
            const editor = {} as IEditorInfo;
            let hasInfo = false;
            row.forEach((value, index) => {
                if (value && positionToNameMap[index]) {
                    editor[positionToNameMap[index] as keyof IEditorInfo] = value;
                    hasInfo = true;
                }
            });
            if (hasInfo && editor.shortName) {
                editorInfoMap[editor.shortName] = editor; //ling_q=>printName:令琪
            }
        }
    }
    return editorInfoMap;
}

function getConfigMapping(values: string[][],log: DebugLog): IGroupAndMainProjectLongToShortNameMapping {
    const configMap: IGroupAndMainProjectLongToShortNameMapping = {
        groupName: '',
        actions: [],
        taskLongToShortNameMapping: {},
        shortProjectNameToProjectId: {} as {
            [key in ActionType]: { project_id: string; }; //populated later after we login to freedcamp
        },
    }
    configMap.groupName = values.find(row => row[0] === 'groupName')?.[1] || '';
    configMap.actions = values.find(row => row[0] === 'operations')?.[1].split(',').map(item => item.trim() as ActionType) || [];
    if (!configMap.groupName) { 
        log.doLog('getConfigMapping: Warning groupName not found in config');
    }
    if (!configMap.actions || configMap.actions.length === 0) {
        log.doLog('getConfigMapping: Warning actions not found in config');
     }
    values.forEach(row => {
        if (row[0] === 'mapping') {
            configMap.taskLongToShortNameMapping[row[1]] = {
                shortName: row[2] as ActionType,
                subTaskOf: row[3] ? row[3] as ActionType : undefined,
            };
        }
    });
    return configMap;
}

type IActionToProjectIdMapping  ={ [key in ActionType]: ProjectAndGroup };
// function getProjectGroupMapping(): IActionToProjectIdMapping {
//     return {
//         '校对': { "project_id": "3696514", "task_group_id": "6825082" },
//         '美编': { "project_id": "3696516", "task_group_id": "6825087" },
//         '发布': { "project_id": "3696243", "task_group_id": "6824401" },
//         '二校': { "project_id": "3708543", "task_group_id": "6856512" },
//     };
// }

function getActionToProjectIdMapping(userData: ICurrentSessionData, groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping): IActionToProjectIdMapping {
    const mapping = {} as IActionToProjectIdMapping;
    userData.data.projects.forEach(proj => {
        const shortProjectName = groupAndMainProjectMapping.taskLongToShortNameMapping[proj.project_name].shortName; ////文字校对 (Editorial and Translation team) : 校对
        if (shortProjectName) {
            mapping[shortProjectName] = { project_id: proj.project_id, task_group_id: '' }
        }
    });
    for (const action of groupAndMainProjectMapping.actions) { //not needed
        const shortProjectName = groupAndMainProjectMapping.taskLongToShortNameMapping[action].shortName; ////文字校对 (Editorial and Translation team) : 校对
        const projectInfo = userData.data.projects.find(proj => proj.project_name === shortProjectName && proj.group_name === groupAndMainProjectMapping.groupName);
    }
    return mapping;
}

async function processOperation(
    opsAndTemplates: OperationAndTemplates,        
    log: DebugLog,
    debug_Prefix: string = ''
): Promise<string[]> {
    const { validOperation, templates, editorInfoMap, groupAndMainProjectMapping } = opsAndTemplates;
    const pr = await login.getProcessor();
    log.doLog('processOperation: got processor with login');
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

        // Check if article is English-only (no Chinese characters)
        const isEnglishOnly = !/[\u4e00-\u9fff]/.test(infos.article);
        
        //const projectGroupMapping = getProjectGroupMapping();
        const projectGroupMapping = getActionToProjectIdMapping(userData, opsAndTemplates.groupAndMainProjectMapping);
        for (const action of opsAndTemplates.groupAndMainProjectMapping.actions) {
            const editor = operation[action];
            const editorLookup = editorInfoMap[editor];
            if (editorLookup) {
                const prettyName = editorLookup.print_name || editorLookup.shortName; 
                if (isEnglishOnly) {
                    infos['editor'] = `${editorLookup.title} ${prettyName}`;
                } else {
                    infos['editor'] = `${prettyName}${editorLookup.title}`;
                }
            }
            const curTemplateActionInfo = templates[action];
            if (curTemplateActionInfo.existingTaskId) {
                log.doLog(`processOperation: skipping action ${action} for file ${fileName} as task ID ${curTemplateActionInfo.existingTaskId} exists`);
                taskIds.push(curTemplateActionInfo.existingTaskId);
                continue;
            }
            let template1 = curTemplateActionInfo.template;
            if (isEnglishOnly && curTemplateActionInfo.templateEnglish) {
                template1 = curTemplateActionInfo.templateEnglish;
            } 
            const projectGrpoup = projectGroupMapping[action];

            const taskRes = await pr.createTask(projectGrpoup, `${debug_Prefix}${fileName}`);
            const taskId = taskRes.id.toString();
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            log.doLog(`processOperation: created task action ${action} ${taskId} for file ${fileName}`);
            await curTemplateActionInfo.taskIdUpdater(taskId);
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
                "priority":2,
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
            log.doLog(`processOperation: prepared post params for task ${taskId} action ${action} with ${JSON.stringify(postParams)}`);
            //due_date
            await pr.doPostAttachment(taskId, postParams);
        }
    }
    return taskIds;
}

export interface DebugLog {
    operations: string[];
    doLog: (msg: string) => void;
}
async function debug_main(lineNumber: number, log:DebugLog, opStr?: string): Promise<boolean> {
    if (opStr === 'del') {
        const ret = await getOpsAndLine(lineNumber, log);
        if (!ret) {
            log.doLog('del: no such line number ' + lineNumber);
            return false;
        }
        const { templates } = ret;
        const pr = await login.getProcessor();
        getActionToProjectIdMapping(await pr.getSessionCurrentData(), ret.groupAndMainProjectMapping);
        for (const action of ret.groupAndMainProjectMapping.actions) {
            const taskId = templates[action].existingTaskId;
            if (taskId) {
                if (taskId === 'done') {
                    const msg = `del: skipping deletion for action ${action} as task ID is marked done`;
                    console.log(msg);
                    log.doLog(msg);
                    continue;
                }
                let msg = `del: deleting task ${taskId} for action ${action}`;
                console.log(msg);
                log.doLog(msg);
                
                await pr.deleteTask(taskId);
                log.doLog(`del: deleted task ${taskId} for action ${action}`);
                await templates[action].taskIdUpdater('');
                log.doLog(`del: cleared task ID in sheet for action ${action}`);
            }
        }
        
        return true;
    }
    const mainResult = await main(lineNumber, log, opStr);
    if (!mainResult) {        
        return false;
    }
    const { ids, ops } = mainResult;
    log.doLog('Created task IDs:'+ JSON.stringify(ids));
    
    //await ops.updateValues('temp!A1', [['forceStringValDEBUGNO_USE,'+ids.join(',')]]);
    
    log.doLog(`Saved ${ids.length} task IDs to temp/taskId.txt`);
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
