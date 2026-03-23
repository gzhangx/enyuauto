
import * as gs from '@gzhangx/googleapi';
import type { ActionType } from './types';
import type { ProjectTaskParams, FreedCampOps, FreedCampProcessor, ICurrentSessionData, IUserInfo, LoginResponse, IProjectTasksResult, FreedCampLoginParams } from './freedcampTypes';
import { getCompleteDateColumnName, getParentTaskIdColumnName, getTaskIdColumnName, type DueDateKeys, type IEditorInfo, type IEditorInfoMap, type IGroupAndMainProjectLongToShortNameMapping, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig, type ISheetDataOps, type ISheetInfoCache, type ISyncFreeCampToSheetData, type OperationInfo, type OperationWithDueDates, type SyncUpdateItem, type Templates } from './opsTypes';
import { findOrCreateExcelFile } from '../src/lib/msGraphConfig';
const mainSheetId = '1zSPJudO0DERn74xV2auIXeNbJxh1apO0tjzB4IrTeQk';

// type DueDateKeys = `${ActionType} Due Date`;
// type TaskIdKeys = `${ActionType} TaskId`;
// interface Operation {
//     '文件': string;
//     '作者': string;
//     '文章名': string;
//     '文章链接': string;
//     '作者电邮': string;
//     '文章类别': string;
//     '校对': string;
//     '美编': string;
//     '发布': string;
//     '二校': string;    
// }

// export function getTaskIdColumnName(action: ActionType): TaskIdKeys {
//     return `${action} TaskId`;
// }

// type OperationWithDueDates = Operation & {
//     [K in DueDateKeys]: string;    
// } & {
//     [k in TaskIdKeys]: string;
// };



// interface OperationInfo {
//     author: string;
//     article: string;
//     link: string;
//     email: string;
//     category: string;
//     editor: string;
// }


// type Templates = {
//     [K in ActionType]: {
//         template: string;
//         templateEnglish: string;
//         taskIdPos: number;
//         //taskIdUpdater: (newTaskId: string, lineNumber: number) => Promise<void>;
//         //getExistingTaskId: (operation: OperationWithDueDates)=>string;
//     }
// };

// interface IGroupAndMainProjectLongToShortNameMapping {
//     freedcampInfo: {
//         username: string;
//         password: string;
//     };
//     groupName: string; //EnYu_2026
//     actions: ActionType[];
//     taskLongToShortNameMapping: {
//         [key: string]: {
//             shortName: ActionType; //文字校对 (Editorial and Translation team) : 校对
//             subTaskOf?: ActionType;
//         }; 
//     };
//     shortProjectNameToProjectId: { //populated later after we login to freedcamp
//         [key in ActionType]: {
//             project_id: string;
//             subTaskOf?: ActionType;
//         }; //'校对': { "project_id": "3696514" }
//     }
// }

// interface IEditorInfoMap { [key: string]: IEditorInfo };


// interface IEditorInfo {
//     title: string;
//     shortName: string;
//     email: string;
//     task: string;
//     print_name: string; //the full chinese name for printing
// }

// interface IOpsConfig {
//     operationList: IOperationWithLineNumber[];
//     groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
//     editorInfoMap: IEditorInfoMap;
//     headers: string[];
//     templates: Templates;
// }


export async function getSheetOps(creds: gs.gsAccount.IServiceAccountCreds, cache?: ISheetInfoCache): Promise<ISheetDataOps> {
    const gsc = await gs.google.gsAccount.getClient(creds);
    const ops = await gsc.getSheetOps(mainSheetId, cache);
    return wrapGoogleSheetOps(ops);
}

/** Wrap a Google Sheets ops object to satisfy ISheetDataOps. */
export function wrapGoogleSheetOps(ops: gs.gsAccount.IGetSheetOpsReturn): ISheetDataOps {
    return {
        readData: (sheet) => ops.readData(sheet),
        autoUpdateValues: async (sheet, values, pos) => { await ops.autoUpdateValues(sheet, values, pos); },
    };
}

function _toExcelCol(n: number): string {
    let result = '';
    while (n > 0) { n--; result = String.fromCharCode(65 + (n % 26)) + result; n = Math.floor(n / 26); }
    return result;
}


/**
 * Create an ISheetDataOps backed by a Microsoft Excel workbook via Graph API.
 * The workbook is resolved by looking up MS_MAIN_EXCEL_FILE_NAME in the
 * General folder (SP_ENYU_GENERAL_FOLDER) of the Shared Documents library on the
 * enyueditors SharePoint site (SP_ENYU_DRIVE_ROOT). The item ID is cached after the first lookup.
 * @param msToken  A valid MS Graph bearer token with Files.ReadWrite scope.
 */
export function createMsExcelDataOps(msToken: string): ISheetDataOps {
    const jsonHeaders = () => ({ Authorization: `Bearer ${msToken}`, 'Content-Type': 'application/json' });

    let cachedItemId: string | null = null;
    let cachedDriveRoot: string | null = null;

    async function getBaseUrl(): Promise<string> {
        if (cachedItemId && cachedDriveRoot) {
            return `${cachedDriveRoot}/items/${cachedItemId}`;
        }
        const { itemId, driveRoot } = await findOrCreateExcelFile(msToken);
        cachedItemId = itemId;
        cachedDriveRoot = driveRoot;
        return `${driveRoot}/items/${itemId}`;
    }

    return {
        async readData(sheetName: string) {
            const baseUrl = await getBaseUrl();
            const res = await fetch(
                `${baseUrl}/workbook/worksheets/${encodeURIComponent(sheetName)}/usedRange`,
                { headers: jsonHeaders() }
            );
            if (!res.ok) {
                const body = await res.json().catch(() => ({}));
                throw new Error((body as any)?.error?.message || res.statusText);
            }
            const data = await res.json();
            return { values: (data.values ?? []) as string[][] };
        },
        async autoUpdateValues(sheetName: string, values: string[][], pos: { row: number; col: number }) {
            const baseUrl = await getBaseUrl();
            // pos.row is 1-indexed data row (header=0), pos.col is 0-indexed
            const physicalRow = pos.row + 1; // spreadsheet row: header at 1, first data at 2
            const startCol = _toExcelCol(pos.col + 1);
            const endCol = _toExcelCol(pos.col + (values[0]?.length ?? 1));
            const endRow = physicalRow + values.length - 1;
            const address = `${startCol}${physicalRow}:${endCol}${endRow}`;
            const res = await fetch(
                `${baseUrl}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${address}')`,
                { method: 'PATCH', headers: jsonHeaders(), body: JSON.stringify({ values }) }
            );
            if (!res.ok) {
                const body = await res.json().catch(() => ({}));
                throw new Error((body as any)?.error?.message || res.statusText);
            }
        },
    };
}

function getExistingTaskId(templateName: ActionType, operation: OperationWithDueDates) {
    return operation[getTaskIdColumnName(templateName)];
}


export async function anyKeyUpdater(ops: ISheetDataOps, opsConfig: IOpsConfig, key: ActionType, item: IOperationWithLineNumber, log: DebugLog) {
    const newValue: string = item[key as keyof typeof item] as string;
    const lineNumber: number = item.itemPositionOnSheet;    
    const index = opsConfig.headers.indexOf(key);
    if (index <= 0) {
        log.doLog(`update column ${key}: ${newValue } at line ${lineNumber} failed due to badk pos ${index}  `);
        return;
    }
    log.doLog(`update column: ${newValue } at line ${lineNumber} for action ${key}  `);
    await ops.autoUpdateValues('main', [[newValue]], {
        row: lineNumber,
        col: index,
    });
}

async function taskIdUpdater(ops: ISheetDataOps, opsConfig: IOpsConfig, key: ActionType, item: IOperationWithLineNumber, log: DebugLog) {
    const newTaskId: string = getExistingTaskId(key, item);
    const lineNumber: number = item.itemPositionOnSheet;
    const index = opsConfig.templates[key].taskIdPos;    
    if (index <= 0) {
        throw new Error(`update taskId: ${newTaskId} at line ${lineNumber} failed due to bad pos ${index} for action ${key}  `);
    }
    log.doLog(`update taskId: ${newTaskId} at line ${lineNumber} index ${index} for action ${key}  `);
    await ops.autoUpdateValues('main', [[newTaskId]], {
        row: lineNumber,
        col: index,
    });
}

export async function completeDateUpdater(ops: ISheetDataOps, opsConfig: IOpsConfig, key: ActionType, item: IOperationWithLineNumber, val: string, log: DebugLog) {
    const lineNumber: number = item.itemPositionOnSheet;
    const index = opsConfig.templates[key].completeDatePos;    
    log.doLog(`update complete date: ${val} at line ${lineNumber} for action ${key}  `);
    await ops.autoUpdateValues('main', [[val]], {
        row: lineNumber,
        col: index,
    });
}

export async function loadMainData(ops: ISheetDataOps) {
    const rawMainData = await ops.readData('main');
    const headers = rawMainData.values[0];
    //const operationListData = await ops.readDataByColumnName('main');
    //const operationList = (operationListData.data || []) as unknown as OperationWithDueDates[];
    const operationList: IOperationWithLineNumberAndParentTaskId[] = rawMainData.values.slice(1).map((row, index) => {
        const obj = {} as IOperationWithLineNumberAndParentTaskId;
        headers.forEach((header, colIndex) => {           
            let val: string = String(row[colIndex] ?? '');
            if (typeof row[colIndex] === 'number') {
                if (header.endsWith('Date')) {
                    // Convert Excel date serial to YYYY-MM-DD
                    // Excel serial 1 = Jan 1 1900; 25569 days between Jan 1 1900 and Unix epoch
                    const date = new Date((row[colIndex] as number - 25569) * 86400 * 1000);
                    const yyyy = date.getUTCFullYear();
                    const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
                    const dd = String(date.getUTCDate()).padStart(2, '0');
                    val = `${yyyy}-${mm}-${dd}`;
                }
            }
            obj[header as keyof OperationWithDueDates] = val.trim();            
        });
        obj.itemPositionOnSheet = index + 1;
        obj.isFinished = false;
        obj.noNeedToCreate = false;
        return obj;
    });
    return { operationList, headers };
}
export async function getOpsAndMainList(ops: ISheetDataOps, log: DebugLog): Promise<IOpsConfig> {
    //const ops = await getSheetOps(token);
    log.doLog('getOpsAndMainList: got sheet ops');
    const { operationList, headers } = await loadMainData(ops);
    //const operationListData = await ops.readDataByColumnName('main');
    //const operationList = (operationListData.data || []) as unknown as OperationWithDueDates[];
    log.doLog('getOpsAndLine: got operation list from main ' + operationList?.length + ' items');
    const configLines = await ops.readData('config');
    log.doLog('getOperationAndTemplates: got config lines ' + configLines.values.length);
    const groupAndMainProjectMapping = getConfigMapping(configLines.values, log);
    
    const editorInfoMap = getEditorInfoMap(configLines.values);

    const templateRows = await ops.readData('templates');
    log.doLog('getOpsAndLine: got template rows');
    const templates = templateRows.values.reduce((acc: Templates, row: string[]) => {
        const [templateName, templateChinese, templateEnglish] = row;
        if (templateName) {
            acc[templateName as ActionType] = {
                template: templateChinese,
                templateEnglish,
                taskIdPos: -1, //headers.values[0].indexOf(templateName),
                completeDatePos: -1,
                //getExistingTaskId: (operation: OperationWithDueDates) => operation[getTaskIdColumnName(templateName as ActionType) as DueDateKeys]
            };
        }
        return acc;
    }, {} as Templates);
    //const headers = await ops.readData('main');
    for (let index = 0; index < headers.length; index++) {
        const item = headers[index];
        for (const key of groupAndMainProjectMapping.actions) {
            if (item === getTaskIdColumnName(key)) {
                log.doLog(`getOpsAndLine: header found ${item} at index ${index}`);
                if (!templates[key]) {
                    throw new Error(`Template for action ${key} not found in templates sheet, but corresponding column ${item} found in main sheet headers. Please fix the templates sheet to include template for ${key}`);
                }
                templates[key].taskIdPos = index;
                //await templates[key].taskIdUpdater('test' + key + "test");                
            } else if (item === getCompleteDateColumnName(key)) {
                templates[key].completeDatePos = index;
            }
        }        
    }

    for (const action of groupAndMainProjectMapping.actions) {
        if (templates[action].taskIdPos === -1) {
            throw new Error(`TaskId column for action ${action} not found in main sheet headers. Expected column name: '${getTaskIdColumnName(action)}'. Please check spelling or add to main sheet header.`);
        }
    }

    //
    return { operationList,groupAndMainProjectMapping,editorInfoMap, headers, templates };
}


export function getOperationFromLineNumber(operationList: IOperationWithLineNumberAndParentTaskId[], lineNumber: number): IOperationWithLineNumberAndParentTaskId | undefined {
    const operation = operationList[lineNumber - 2];
    return operation;
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
            [key in ActionType]: {
                project_id: string;
                subTaskOf?: ActionType;
            }; //populated later after we login to freedcamp
        },
        actionExcludes: {},
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
                shortName: row[2].trim() as ActionType,
                subTaskOfFromSheetConfig: row[3] ? row[3].trim() as ActionType : undefined,
                isTaskEnabledForEnglishFromSheetConfig: row[4]?.trim() === 'N' ? 'N' : '',
            };
        }
        if (row[0] === 'actionExcludes') {
            const action = row[1].trim() as ActionType;
            const excludes = row[2].split(',').map(item => item.trim() as ActionType);
            configMap.actionExcludes[action] = excludes;
        }
    });
    return configMap;
}

type IActionToProjectIdMapping  ={ [key in ActionType]: ProjectTaskParams };
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
        const projectInfo = groupAndMainProjectMapping.taskLongToShortNameMapping[proj.project_name.trim()]; ////文字校对 (Editorial and Translation team) : 校对
        if (projectInfo && projectInfo.shortName) {
            mapping[projectInfo.shortName] = {
                project_id: proj.project_id,
                priority: 2,
            }

            groupAndMainProjectMapping.shortProjectNameToProjectId[projectInfo.shortName] = {
                project_id: proj.project_id,
                subTaskOf: projectInfo.subTaskOfFromSheetConfig,
                isTaskEnabledForEnglish: projectInfo.isTaskEnabledForEnglishFromSheetConfig !== 'N',
            };
        }
    });
    //for (const action of groupAndMainProjectMapping.actions) { //not needed
    //    const shortProjectName = groupAndMainProjectMapping.taskLongToShortNameMapping[action].shortName; ////文字校对 (Editorial and Translation team) : 校对
    //    const projectInfo = userData.data.projects.find(proj => proj.project_name === shortProjectName && proj.group_name === groupAndMainProjectMapping.groupName);
    //}
    
    // Validate that all required mappings are present
    const missingKeys: string[] = [];
    for (const [longName, projectInfo] of Object.entries(groupAndMainProjectMapping.taskLongToShortNameMapping)) {
        if (projectInfo.shortName && !mapping[projectInfo.shortName]) {
            missingKeys.push(`${longName} (shortName: ${projectInfo.shortName})`);
        }
    }
    
    if (missingKeys.length > 0) {
        throw new Error(`Missing project mappings for the following keys: ${missingKeys.join(', ')}`);
    }
    
    return mapping;
}

export interface FreeCampAndUpdateOperations {
    getFreedCampToken: (loginParams: FreedCampLoginParams) => Promise<LoginResponse>;
    getFreedCampProcessor: (loginToken: LoginResponse) => FreedCampProcessor;
}

export function getFreeCampAndUpdateOperations(freedCampOps: FreedCampOps): FreeCampAndUpdateOperations {
    async function getFreedCampToken(loginParams: FreedCampLoginParams) {
        const loginToken = await freedCampOps.login(loginParams);
        return loginToken;
    }
    function getFreedCampProcessor(loginToken: LoginResponse) {
        const pr = freedCampOps.getProcessor(loginToken);
        return pr;
    }
    return {
        getFreedCampToken,
        getFreedCampProcessor,
    }
}

export interface ICombinedOpsAndFreeCampData {
    opsConfig: IOpsConfig;
    userData: ICurrentSessionData;
    userNameToInfoMap: { [shortName: string]: IUserInfo };
    userIdToInfoMap: { [userId: string]: IUserInfo };
    projectGroupMapping: IActionToProjectIdMapping;
    loginToken: LoginResponse;
    freedCampTasksByAction: { [key in ActionType]: IProjectTasksResult };
}
export async function combineOpsConfigWithFreedCampData(opsConfig: IOpsConfig, freedCampOps: FreeCampAndUpdateOperations, freedCampLoginParams: FreedCampLoginParams, log: DebugLog)
:Promise<ICombinedOpsAndFreeCampData>
{
    const loginToken = await freedCampOps.getFreedCampToken(freedCampLoginParams);
    const pr = freedCampOps.getFreedCampProcessor(loginToken);
    log.doLog('processOperation: got processor with login');
    
    const userData = await pr.getSessionCurrentData();
    const userIdToInfoMap: { [userId: string]: IUserInfo } = {};
    const userNameToInfoMap = userData.data.users.reduce((acc, user) => {
        acc[user.full_name] = user;
        userIdToInfoMap[user.user_id] = user;
        return acc;
    }, {} as { [key: string]: IUserInfo });
    const projectGroupMapping = getActionToProjectIdMapping(userData, opsConfig.groupAndMainProjectMapping);

    return {
        loginToken,
        opsConfig,
        userData,
        userNameToInfoMap,
        userIdToInfoMap,
        projectGroupMapping,
        freedCampTasksByAction: {} as { [key in ActionType]: IProjectTasksResult },
    }
}


function getOperationInfo(combined: ICombinedOpsAndFreeCampData,    
    operation: IOperationWithLineNumberAndParentTaskId, log: DebugLog,) {
    const article = operation['文章名'];
    //const fileName = operation['文件'];
    const infos: OperationInfo = {        
            author: operation['作者'],
            article,
            link: operation['文章链接'],
            email: operation['作者电邮'],
        category: operation['文章类别'],
            mainFolder: operation['mainFolder'],
            //校对: operation['校对'],
            editor: '',            
            //isEnglishOnly: !/[\u4e00-\u9fff]/.test(fileName)?'Y':'N', // Check if article contains Chinese characters
        };

        // Check if article is English-only (no Chinese characters)
    const groupAndMainProjectMapping = combined.opsConfig.groupAndMainProjectMapping;
    for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
        const actionConfig = groupAndMainProjectMapping.shortProjectNameToProjectId[action];
        const editor = operation[action];
        const editorLookup = combined.opsConfig.editorInfoMap[editor];

        const isEnglishOnly =isEditorEnglish(editorLookup);
        if (actionConfig.isTaskEnabledForEnglish === false && isEnglishOnly) {
            const fileName = operation['文件'];
            log.doLog(`processOperation: skipping action ${action} for file ${fileName} as it is disabled from sheet config for English-only article`);
            operation[getTaskIdColumnName(action)] = 'TaskIsEnglishAndIsDisabledForEnglish';
            continue;
        }
    }
    
    return infos;
}



export function formatLocalDateYyyyMmDd(unixSeconds: number): string {
	const date = new Date(unixSeconds * 1000);
	return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function buildSyncUpdates(
	p: IOperationWithLineNumberAndParentTaskId,
	combined: ICombinedOpsAndFreeCampData,
): ISyncFreeCampToSheetData | undefined {
	const parts: string[] = [];
	const updates: SyncUpdateItem[] = [];

	for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
		const freedCampItem = p[`${action} FreeCamp Item`];
		if (!freedCampItem) {
			parts.push(`Action ${action}: No FreedCamp item, skipped`);
			continue;
		}

		const columnMapping: { [sheetCol: string]: keyof typeof freedCampItem } = {
            [action]: 'assigned_to_id',
            [`${action} TaskId`]: 'id',
			[`${action} Due Date`]: 'due_ts',
			[`${action} Complete Date`]: 'completed_ts',
		};

		Object.entries(columnMapping).forEach(([sheetCol, freedCampKey]) => {
			const sheetVal = p[sheetCol as keyof typeof p];
			let fcVal = freedCampItem[freedCampKey] as string;
			let compareValues: string[] = [];

            if (freedCampKey !== 'id') {
                if (freedCampKey === 'assigned_to_id') {
                    const userInfo = combined.userIdToInfoMap[fcVal];
                    console.log(`Mapping user ID ${fcVal} to name:`, userInfo, freedCampItem, freedCampKey);
                    if (userInfo) {
                        fcVal = userInfo.full_name;
                    }
                    if (!fcVal || fcVal === '0') {
                        return;
                    }
                } else if (fcVal) {
                    const ts = Number(fcVal);
                    if (!Number.isNaN(ts)) {
                        const date = new Date(ts * 1000);
                        const yyyyMmDd = formatLocalDateYyyyMmDd(ts);
                        const mDyYyyy = `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
                        compareValues = [yyyyMmDd, mDyYyyy];
                        fcVal = yyyyMmDd;
                    }
                }
            }

			const colIdx = combined.opsConfig.headers.indexOf(sheetCol);
			if (colIdx === -1) {
				parts.push(`Column ${sheetCol} not found in sheet headers`);
				return;
			}

			const sheetValStr = String(sheetVal);
			const matchesSheet = compareValues.length > 0
				? compareValues.includes(sheetValStr)
				: String(fcVal) === sheetValStr;

			if (fcVal !== undefined && fcVal !== null && !matchesSheet) {
				const displayVal = compareValues.length > 0 ? `${compareValues[0]} / ${compareValues[1]}` : String(fcVal);
				parts.push(
					`Action ${action}: SheetCol "${sheetCol}" (col ${colIdx + 1}, line ${p.itemPositionOnSheet}) will update to "${displayVal}" (was "${sheetVal}")`
				);
				updates.push({ sheetCol, value: fcVal });
			}
		});
	}

    if (updates.length === 0) return undefined;
	return { parts, updates };
}

export function updateDoneParentIds(combinedOpsAndData: ICombinedOpsAndFreeCampData, projectList: IOperationWithLineNumberAndParentTaskId[],log: DebugLog) {
    const opsConfig = combinedOpsAndData.opsConfig;
    const actionToTitleToProjectAndTaskIdMap: {
      [action in ActionType]?: {
        [title: string]: {
          projectId: string;
          taskId: string;
          task: IOperationWithLineNumberAndParentTaskId;
        };
      };
    } = {};
    console.log('Checking for done sub-tasks to mark main tasks as done...');
      opsConfig.groupAndMainProjectMapping.actions.forEach(action => {
        const actionInfo = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
        if (actionInfo && actionInfo.subTaskOf) {
          const subTaskOfAction = actionInfo.subTaskOf;
          projectList.forEach(item => {
            if (item[getTaskIdColumnName(subTaskOfAction)] === 'done' && !item[getTaskIdColumnName(action)]) {
              console.log(`EARNNN ${item.文件} ${item.文章名} ${action} as done for line ${item.itemPositionOnSheet} because ${subTaskOfAction} is done`);
              const prj = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[subTaskOfAction];
              if (prj) {
                console.log(`Debugrm parent project ${subTaskOfAction} project id ${prj.project_id} to check if all sub-tasks are done`);
                if (!actionToTitleToProjectAndTaskIdMap[subTaskOfAction]) actionToTitleToProjectAndTaskIdMap[subTaskOfAction] = {};                
                actionToTitleToProjectAndTaskIdMap[subTaskOfAction]![item.文件] = {
                  projectId: prj.project_id,
                  taskId: '',
                  task: item,
                };
              }
            }
          });
        }
      });
    
    let updated = false;
    if (combinedOpsAndData)
    {
      //const pr = freeCampOpsWithCache.getFreedCampProcessor(loginToken);
      for (const action of opsConfig.groupAndMainProjectMapping.actions) {
        const actionInfo = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
        if (actionInfo) {
          //setIsLoading(prev => ({ ...prev, freeCampLoading: `Loading FreeCamp tasks for ${action}...` }));
          const prjs = combinedOpsAndData.freedCampTasksByAction[action];
          for (const itm of projectList) {
            let freedCampTsk = prjs.tasks.find(tsk => tsk.title === itm.文件);
            if (freedCampTsk) {
              itm[`${action} FreeCamp Item`] = freedCampTsk;
            } else {
                const freedCampTsks = prjs.tasks.filter(tsk => tsk.title?.includes(itm.文件));
                if (freedCampTsks.length !== 0) {
                    if (freedCampTsks.length > 1) {
                        log.doLog(`Warning: multiple FreedCamp tasks found for ${itm.文件} in action ${action}, assigning the first one. Task IDs: ${freedCampTsks.map(t => t.id).join(', ')}`, true);
                    } else {
                        itm[`${action} FreeCamp Item`] = freedCampTsks[0];
                    }
                }
            }
          }
        }
      }

        //const subTaskWithParentDone = actionToTitleToProjectAndTaskIdMap[action];      
      // for (const actionName of Object.keys(actionToTitleToProjectAndTaskIdMap)) {
      //   const cnt = actionToTitleToProjectAndTaskIdMap[actionName as ActionType] || {};
      //   for (const title of Object.keys(cnt)) {
      //     const tskAndPrj = cnt[title];
      //     const prjs = await pr.getTasksForProjects(tskAndPrj.projectId, 1);
      //     const tsk = prjs.tasks.find(tsk => tsk.title === title);
      //     if (tsk) {
      //       tskAndPrj.taskId = tsk.id || '';
      //       tskAndPrj.task[`${actionName as ActionType} ParentTaskId`] = tskAndPrj.taskId;
      //       console.log('debugremove setting task id for', title, 'to', tsk.id);
      //       updated = true;
      //     }
      //   }        
      // }
    }

    for (const item of projectList) {
        getOperationInfo(combinedOpsAndData, item, log);
        const syncFreeCompToSheetParts = buildSyncUpdates(item, combinedOpsAndData);
        item.syncFreeCampToSheetData = syncFreeCompToSheetParts;
        const publishItem = item[`发布 FreeCamp Item`];
        let publishCompleted = false;
        if (publishItem) {
            publishCompleted = publishItem.completed_ts !== null && publishItem.completed_ts !== undefined;
        }
        const taskDone = item['发布'] === 'done';
        item.isFinished = (publishCompleted || taskDone) && !item.syncFreeCampToSheetData;
        item.noNeedToCreate = publishCompleted || taskDone;
    }
    if (updated) {
      //setProjectList([...projectList]);
    }
    //setIsLoading(prev => ({ ...prev, freeCampLoading: '' }));
}

function getEditorNameForAction(
    editor: string,
    combined: ICombinedOpsAndFreeCampData,
): string {
    //const editor = operation[action];
    const editorLookup = combined.opsConfig.editorInfoMap[editor];
    if (!editorLookup) {
        return '';
    }

    const prettyName = editorLookup.print_name || editorLookup.shortName;
    const isEnglishOnly = isEditorEnglish(editorLookup);
    return isEnglishOnly
        ? `${editorLookup.title} ${prettyName}`
        : `${prettyName}${editorLookup.title}`;
}

function isEditorEnglish(editorLookup: IEditorInfo): boolean {
    if (!editorLookup) return false;
    const normalizedTitle = editorLookup.title.toLowerCase();
    return normalizedTitle === 'brother' || normalizedTitle === 'sister';
}

export interface IOneDriveDirInfo {
    name: string;
    webUrl: string;
}


interface ActionsToPerformInfo {
    action: ActionType;
    editor: string;
     editorName: string;
     template: string;
}
/**
 * Returns the list of actions that still need to be performed for a given operation.
 * Mirrors the filtering logic in processOperation:
 *   - Skips actions whose task ID is already set.
 *   - Skips actions disabled for English-only articles.
 *   - When an action has its AI flag set (e.g. `校对 AI` = 'Y') AND a corresponding AI template
 *     exists, it is treated as an AI action. In that case actionExcludes is checked using the AI
 *     action name (e.g. `校对 AI`) to exclude other actions from the result.
 */
export function getActionsToPerform(
    operation: IOperationWithLineNumberAndParentTaskId,
    combined: ICombinedOpsAndFreeCampData,
    fileName: string, //fileName, for log only
    log: DebugLog,
): ActionsToPerformInfo[] {
    const { groupAndMainProjectMapping, templates } = combined.opsConfig;

    //real actions, replacing name with AI if found. however since config is still none-ai, this is only used to do exclusions.
    const allRealActions: ActionType[] = groupAndMainProjectMapping.actions.map(action => {
        const AIActionName = `${action} AI` as ActionType;
        const isAIAction = (operation[AIActionName] || '').toUpperCase() === 'Y';
        if (isAIAction && templates[AIActionName]) {
            return AIActionName;
        } else {
            return action;
        }
    });    

    const excludedActions = new Set<ActionType>();
    for (const action of allRealActions) {
        const excludes = groupAndMainProjectMapping.actionExcludes[action];
        if (excludes) {
            for (const excluded of excludes) {
                excludedActions.add(excluded);
            }
        }
    }

    const actionsToPerform: ActionsToPerformInfo[] = [];
    const taskIds: string[] = [];
    for (const action of groupAndMainProjectMapping.actions) {
        if (excludedActions.has(action)) {
            log.doLog(`getActionsToPerform: excluding action ${action} due to actionExcludes`);
            continue;
        }

        const actionConfig = groupAndMainProjectMapping.shortProjectNameToProjectId[action];
        if (!actionConfig) {
            throw new Error(`Configuration for action ${action} not found in groupAndMainProjectMapping.shortProjectNameToProjectId`);
        }
        const editor = operation[action];
        const editorLookup = combined.opsConfig.editorInfoMap[editor];
        const isEnglishOnly = isEditorEnglish(editorLookup);
        if (actionConfig.isTaskEnabledForEnglish === false && isEnglishOnly) {
            log.doLog(`processOperation: skipping action ${action} for file ${fileName} as it is disabled from sheet config for English-only article`);
            operation[getTaskIdColumnName(action)] = 'TaskIsEnglishAndIsDisabledForEnglish';
            continue;
        }
        const editorName = getEditorNameForAction(editor, combined) || 'EDITOR NOTSET';

        const AIActionName = `${action} AI` as ActionType;
        const isAIAction = (operation[AIActionName] || '').toUpperCase() === 'Y';
        const regularAction = templates[action];
        const curTemplateActionInfo = isAIAction ? templates[AIActionName] || templates[action] : regularAction;
        if (isAIAction && curTemplateActionInfo === regularAction) {
            log.doLog(`processOperation: AI action ${AIActionName} is enabled for file ${fileName} but template for it is not found, falling back to regular action ${action} template`, true);
        }
        const existingTaskId = getExistingTaskId(action, operation);
        if (existingTaskId) {
            log.doLog(`processOperation: skipping action ${action} for file ${fileName} as task ID ${existingTaskId} exists`);
            taskIds.push(existingTaskId);
            continue;
        }
        let template1 = curTemplateActionInfo.template;
        if (isEnglishOnly && curTemplateActionInfo.templateEnglish) {
            template1 = curTemplateActionInfo.templateEnglish;
        }
        //const link = action === '校对' ? infos.link : undefined;

        actionsToPerform.push({
            action,
            editorName,
            editor,
            template: template1,
        });
    }
    return actionsToPerform;
}

export async function processOperation(
    ops: ISheetDataOps,
    freedCampOps: FreeCampAndUpdateOperations,
    combined: ICombinedOpsAndFreeCampData,    
    operation: IOperationWithLineNumberAndParentTaskId,    
    oneDriveDirInfos: IOneDriveDirInfo[],
    log: DebugLog,
    debug_Prefix: string = ''
): Promise<string[]> {
    const { groupAndMainProjectMapping } = combined.opsConfig;    
    log.doLog('processOperation: got processor with login');
    
    const taskIds = [];
    
    const pr = freedCampOps.getFreedCampProcessor(combined.loginToken);
    {
        const fileName = operation['文件'];
        const infos: OperationInfo = getOperationInfo(combined, operation, log);

        // Check if article is English-only (no Chinese characters)
        
        const actionsToPerform = getActionsToPerform(operation, combined, fileName, log);
        //const projectGroupMapping = getProjectGroupMapping();        
        //for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
        for (const actionP of actionsToPerform) {
            const action = actionP.action;
            const actionConfig = groupAndMainProjectMapping.shortProjectNameToProjectId[action];           
            infos.editor = actionP.editorName;
            //const editor = operation[action];
            //const editorLookup = combined.opsConfig.editorInfoMap[editor];
            //const isEnglishOnly = isEditorEnglish(editorLookup);
            // if (actionConfig.isTaskEnabledForEnglish === false && isEnglishOnly) {
            //    log.doLog(`processOperation: skipping action ${action} for file ${fileName} as it is disabled from sheet config for English-only article`);
            //    operation[getTaskIdColumnName(action)] = 'TaskIsEnglishAndIsDisabledForEnglish';
            //    continue;
            //}
            //infos.editor = getEditorNameForAction(editor, combined) || 'EDITOR NOTSET';

            //const AIActionName = `${action} AI` as ActionType;
            //const isAIAction = (operation[AIActionName] || '').toUpperCase() === 'Y';
            //const regularAction = templates[action];            
            //const curTemplateActionInfo = isAIAction ? templates[AIActionName] || templates[action] : regularAction;
            //if (isAIAction && curTemplateActionInfo === regularAction) {
            //    log.doLog(`processOperation: AI action ${AIActionName} is enabled for file ${fileName} but template for it is not found, falling back to regular action ${action} template`, true);
            //}
            //const existingTaskId = getExistingTaskId(action, operation);
            //if (existingTaskId) {
            //    log.doLog(`processOperation: skipping action ${action} for file ${fileName} as task ID ${existingTaskId} exists`);
            //    taskIds.push(existingTaskId);
            //    continue;
            //}
            //let template1 = curTemplateActionInfo.template;
            //if (isEnglishOnly && curTemplateActionInfo.templateEnglish) {
            //    template1 = curTemplateActionInfo.templateEnglish;
            //} 
            let template1 = actionP.template;
            const editor = operation[action];
            const link = action === '校对' ? infos.link : undefined;            
            
            // if oneDriveDirInfos exists, find all matches in template that matches  {media_link:XXXX}, and if XXXX matches any of the oneDriveDirInfos.name, replace {media_link:XXXX} 
            // with the corresponding oneDriveDirInfos.webUrl
            if (oneDriveDirInfos && oneDriveDirInfos.length > 0) {
                template1 = template1.replace(/\{media_link:([^}]+)\}/g, (match, name) => {
                    const found = oneDriveDirInfos.find(info => info.name === name);
                    return found ? found.webUrl : match;
                });
            }
            for (const replaceItem of ['editor', 'author', 'email', 'article', 'category', 'slug']) {
                if (replaceItem === 'article' && link) continue;
                template1 = template1.replaceAll(`{${replaceItem}}`, infos[replaceItem as keyof OperationInfo]);
            }
            if (link) {
                template1 = template1.replaceAll('{article}', `<a href="${infos['link']}">${infos['article']}</a>`);
            }


            let projectGrpoup = combined.projectGroupMapping[action];
            const subTaskOf = actionConfig.subTaskOf;
            let taskTitle = `${debug_Prefix}${fileName}`
            if (subTaskOf) {
                let h_parent_id = getExistingTaskId(subTaskOf, operation);
                if (!h_parent_id || h_parent_id === 'done') {
                    h_parent_id = operation[getParentTaskIdColumnName(subTaskOf)] as string;
                    if (!h_parent_id) {
                        log.doLog(`processOperation: Warning no parent task ID found for subtask action ${action} with parent action ${subTaskOf} for file ${fileName}`);
                        throw new Error(`No parent task ID found for subtask action ${action} with parent action ${subTaskOf}`);
                    }

                }
                projectGrpoup = { ...combined.projectGroupMapping[subTaskOf] };
                projectGrpoup.h_parent_id = h_parent_id;
                projectGrpoup.description = template1;
                taskTitle = action;
            }
            const taskRes = await pr.createTask(projectGrpoup, taskTitle);
            const taskId = taskRes.id.toString();
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            log.doLog(`processOperation: created task action ${action} ${taskId} for file ${fileName}`);
            operation[getTaskIdColumnName(action)] = taskId;
            await taskIdUpdater(ops, combined.opsConfig, action, operation, log);
            taskIds.push(taskId);
            
            
            const postParams: ProjectTaskParams = {
                description: template1,
                "priority":2,
            };
            const dueDateKey = `${action} Due Date` as DueDateKeys;
            const due_date = operation[dueDateKey];

            if (due_date) {
                // Convert due_date string to Unix timestamp (seconds)
                const dueDateObj = new Date(due_date);
                if (!isNaN(dueDateObj.getTime())) {
                    postParams.due_ts = Math.floor(dueDateObj.getTime() / 1000);
                } 
            }
            if (editor) {
                const userInfo = combined.userNameToInfoMap[editor];
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
            if (!subTaskOf) {
                await pr.doPostAttachment(taskId, postParams);
            }
        }
    }
    return taskIds;
}

export interface DebugLog {
    doLog: (msg: string, critical?: boolean) => void;
}


export async function deleteItemActionTask( ops: ISheetDataOps,
    freedCampOps: FreeCampAndUpdateOperations, 
    freedCampLoginParams: FreedCampLoginParams,
    opsAndTemplates: IOpsConfig,
    operation: IOperationWithLineNumber,    
    action: ActionType,
    log: DebugLog,) {
    
    const loginToken = await freedCampOps.getFreedCampToken(freedCampLoginParams);
    const pr = freedCampOps.getFreedCampProcessor(loginToken);
    const existingTaskId = getExistingTaskId(action, operation);
    if (existingTaskId) {
        if (existingTaskId === 'done') {
            const msg = `del: skipping deletion for action ${action} as task ID is marked done`;
            console.log(msg);
            log.doLog(msg);
            return msg;
        }
        let msg = `del: deleting task ${existingTaskId} for action ${action}`;
        console.log(msg);
        log.doLog(msg);
                
        await pr.deleteTask(existingTaskId);
        log.doLog(`del: deleted task ${existingTaskId} for action ${action}`);
        operation[getTaskIdColumnName(action)] = '';
        await taskIdUpdater(ops, opsAndTemplates, action, operation, log);
        log.doLog(`del: cleared task ID in sheet for action ${action}`);
    }
}


export async function main(ops: ISheetDataOps, freedCampOps: FreedCampOps, freedCampLoginParams: FreedCampLoginParams, lineNumber: number, log: DebugLog, opStr?: string) {
    const opsAndTemplates = await getOpsAndMainList(ops, log);
    if (!opsAndTemplates) {        
        return undefined;
    }
    const prefix = opStr || '';    
    const fops = getFreeCampAndUpdateOperations(freedCampOps);
    const validOperation = getOperationFromLineNumber(opsAndTemplates.operationList, lineNumber)!;
    if (!validOperation) {
        log.doLog('processOperation: no such line number ' + lineNumber);
        return undefined;
    }
    const combined = await combineOpsConfigWithFreedCampData(opsAndTemplates, fops, freedCampLoginParams, log);
    const ids = await processOperation(ops, fops, combined, validOperation, [], log, prefix);
    return {
        ids, 
        ...opsAndTemplates,
    }
}
