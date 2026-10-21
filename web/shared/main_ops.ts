
import * as gs from '@gzhangx/googleapi';
import type { ActionType } from './types';
import type { ProjectTaskParams, FreedCampOps, FreedCampProcessor, ICurrentSessionData, IUserInfo, LoginResponse, IProjectTasksResult, FreedCampLoginParams } from './freedcampTypes';
import { getCompleteDateColumnName, getTaskIdColumnName, type DueDateKeys, type IEditorInfo, type IGroupAndMainProjectLongToShortNameMapping, type IOnLeaveData, type IOperationWithLineNumber, type IOperationWithLineNumberAndParentTaskId, type IOpsConfig, type ISheetDataOps, type ISheetInfoCache, type ISyncFreeCampToSheetData, type OperationInfo, type OperationWithDueDates, type ProjectActionMappingConfig, type SyncUpdateItem, type Templates } from './opsTypes';
import { copySharePointFile, findOrCreateExcelFile, joinSharePointPath, type SharePointFileRef } from '../src/lib/msGraphConfig';
import { detectArticleLanguage, findRelatedEnyuArticles, formatRelatedArticlesHtml } from '../src/lib/enyuRelatedArticles';
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
//     fullNamePN: string; //the full chinese name for printing
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
        readRange: (sheet, address) => ops.readData(`${sheet}!${address}`),
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
            const data = await res.json() as { values?: string[][] };
            return { values: (data.values ?? [])};
        },
        async readRange(sheetName: string, address: string) {
            const baseUrl = await getBaseUrl();
            const res = await fetch(
                `${baseUrl}/workbook/worksheets('${encodeURIComponent(sheetName)}')/range(address='${address}')`,
                { headers: jsonHeaders() }
            );
            if (!res.ok) {
                const body = await res.json().catch(() => ({}));
                throw new Error((body as any)?.error?.message || res.statusText);
            }
            const data = await res.json() as { values?: string[][] };
            return { values: data.values ?? [] };
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

/** Writes the main sheet `done` column: `Y` hides the row when "Show all" is off (`isFinished`). */
export async function updateDoneColumn(
    ops: ISheetDataOps,
    opsConfig: IOpsConfig,
    item: IOperationWithLineNumber,
    markDone: boolean,
    log: DebugLog,
): Promise<void> {
    const lineNumber = item.itemPositionOnSheet;
    const doneHeader = 'done';
    const index = opsConfig.headers.indexOf(doneHeader);
    if (index < 0) {
        throw new Error(`Main sheet header row must include a column named "${doneHeader}".`);
    }
    const value = markDone ? 'Y' : '';
    log.doLog(`update done column: ${value || '(empty)'} at line ${lineNumber}`);
    await ops.autoUpdateValues('main', [[value]], {
        row: lineNumber,
        col: index,
    });
}


const NEW_PROJECT_COLUMNS = ['文件', 'mainFolder', '作者', '称呼', '作者电邮', '文章类别', '文章名', '文章链接', 'slug'] as const;

/** Appends one main-sheet row after the last row that already has a 文件 or 文章名. Returns the 1-based data row written. */
export async function appendMainSheetRow(
    ops: ISheetDataOps,
    headers: string[],
    existingRows: { itemPositionOnSheet: number; 文件?: string; 文章名?: string }[],
    values: Record<string, string>,
    log: DebugLog,
): Promise<number> {
    const missing = NEW_PROJECT_COLUMNS.filter(name => !headers.includes(name));
    if (missing.length > 0) {
        throw new Error(`Main sheet is missing columns: ${missing.join(', ')}`);
    }
    let last = 0;
    for (const row of existingRows) {
        if ((row.文件 || '').trim() || (row.文章名 || '').trim()) {
            last = Math.max(last, row.itemPositionOnSheet);
        }
    }
    const lineNumber = last + 1;
    const rowValues = headers.map(header => values[header] ?? '');
    log.doLog(`append main row at line ${lineNumber}: ${values['文件'] || ''}`);
    await ops.autoUpdateValues('main', [rowValues], { row: lineNumber, col: 0 });
    return lineNumber;
}

export async function updateAnyColumn(
    ops: ISheetDataOps,
    opsConfig: IOpsConfig,
    item: IOperationWithLineNumber,
    columnName: string,
    value: string,
    log: DebugLog,
): Promise<void> {
    const lineNumber = item.itemPositionOnSheet;
    const index = opsConfig.headers.indexOf(columnName);
    if (index < 0) {
        throw new Error(`Main sheet header row must include a column named "${columnName}".`);
    }    
    log.doLog(`update generic column:  ${columnName}===>${value || '(empty)'} at line ${lineNumber}`);
    await ops.autoUpdateValues('main', [[value]], {
        row: lineNumber,
        col: index,
    });
}

export async function loadMainData(ops: ISheetDataOps) {
    const rawMainData = await ops.readData('main');
    const headers = rawMainData.values[0].map(h => h.toLowerCase() === 'done' ? 'done' : h);
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

function excelToJSDate(excelSerial: number) {
    // Excel erroneously treats 1900 as a leap year, so the base offset is 25569 days to Unix epoch
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    return new Date((excelSerial - 25569) * MS_PER_DAY);
}

async function getOnLeaveData(ops: ISheetDataOps, editorInfoAry: IEditorInfo[], log: DebugLog): Promise<IOnLeaveData[]> {
    //const ops = await getSheetOps(token);
    log.doLog('getOpsAndMainList: got onLeave data');
    const onLeaveData = await ops.readData('OnLeave');
    const onLeaves: IOnLeaveData[] = onLeaveData.values.map((v, idx) => {        
        const name = v[0];
        if (name) {
            const matched = editorInfoAry.find(e => e.fullNamePN === name || e.shortNameOnFreedCamp === name);
            if (!matched) {
                if (idx === 0) return;
                log.doLog(`Warning, onLeave name ${name} does not match anytthing!!!!!`)
                return;
            }
            const ret: IOnLeaveData = {
                name: matched.fullNamePN,
                startDate: excelToJSDate(v[1] as unknown as number),
                endDate: excelToJSDate(v[2] as unknown as number),
            }
            if (ret.endDate < ret.startDate) {
                log.doLog(`Invalid date range for ${ret.name}: End date cannot be before start date. ${ret.startDate} ${ret.endDate}`);
                // Handle the error here (e.g., throw an error, skip the record, or set a fallback)
            }   
            //console.log(name, matched.fullNamePN, v[1], v[2], ret.startDate.toISOString(), ret.endDate.toISOString());
            return ret;
        }
    }).filter(x => x) as IOnLeaveData[];        
    return onLeaves;

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
    
    const editorInfoAry = getEditorInfoAry(configLines.values);

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
    const onLeaveData = await getOnLeaveData(ops, editorInfoAry, log);
    const { categories, titles } = await getCategorySheetLists(ops, log);
    return { operationList,groupAndMainProjectMapping,editorInfoAry, headers, templates, onLeaveData, categories, titles };
}

const CATEGORY_SHEET_NAMES = ['categories', 'catgegories', 'categogires'];

function normalizeListLabel(value: string): string {
    return value.normalize('NFKC').replace(/[\u200b\u200c\u200d\ufeff]/g, '').trim().replace(/\s+/g, ' ');
}

function uniqueColumn(values: string[][], column: number, foldTraditionalShi: boolean): string[] {
    const result: string[] = [];
    const seen = new Set<string>();
    for (const row of values) {
        let name = normalizeListLabel(String(row?.[column] ?? ''));
        if (foldTraditionalShi) {
            name = name.replace(/師/g, '师');
        }
        const key = name.toLowerCase();
        if (!name || seen.has(key)) {
            continue;
        }
        seen.add(key);
        result.push(name);
    }
    return result;
}

/** Categories in A2:A11 and 称呼 in E2:E11. Sheet order is kept. */
async function getCategorySheetLists(ops: ISheetDataOps, log: DebugLog): Promise<{ categories: string[]; titles: string[] }> {
    let values: string[][] | null = null;
    for (const sheetName of CATEGORY_SHEET_NAMES) {
        try {
            const data = await ops.readRange(sheetName, 'A2:E11');
            values = data.values;
            log.doLog(`getCategorySheetLists: read ${values.length} rows from ${sheetName}!A2:E11`);
            break;
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            log.doLog(`getCategorySheetLists: ${sheetName} not read (${message})`);
        }
    }
    if (!values) {
        return { categories: [], titles: [] };
    }
    return {
        categories: uniqueColumn(values, 0, false),
        titles: uniqueColumn(values, 4, true),
    };
}


export function getOperationFromLineNumber(operationList: IOperationWithLineNumberAndParentTaskId[], lineNumber: number): IOperationWithLineNumberAndParentTaskId | undefined {
    const operation = operationList[lineNumber - 2];
    return operation;
}





function getEditorInfoAry(values: string[][]): IEditorInfo[] {
    const editorInfoAry: IEditorInfo[] = [];
    const positionToNameMap: { [key: number]: string } = {};
    (values.find(row => row[0] === 'EditorNameColumn') || []).forEach((name, index) => {
        switch (name) {
            case 'Full Name':
                positionToNameMap[index] = 'fullNamePN';
                break;
            case 'Group':
                positionToNameMap[index] = 'group'; //校对/二校 etc
                break;            
            case 'Name on FreedCamp':
                positionToNameMap[index] = 'shortNameOnFreedCamp';
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
            if (hasInfo && editor.shortNameOnFreedCamp) {
                //editorInfoMap[editor.shortNameOnFreedCamp] = editor; //ling_q=>printName:令琪
                editorInfoAry.push(editor);
            }
        }
    }
    //Name on FreedCamp (ric.m etc)->{fullNamePN, group, shortName, title}
    return editorInfoAry;
}

function getConfigMapping(values: string[][],log: DebugLog): IGroupAndMainProjectLongToShortNameMapping {
    const configMap: IGroupAndMainProjectLongToShortNameMapping = {
        groupName: '',
        actions: [],
        taskLongToShortNameMapping: {},
        shortProjectNameToProjectId: {} as {
            [key in ActionType]: {
                project_id: string;
                mappingConfig: ProjectActionMappingConfig;
            }; //populated later after we login to freedcamp
        },
        actionExcludes: {},
        autoAssign: [],
    }
    configMap.groupName = values.find(row => row[0] === 'groupName')?.[1] || '';
    configMap.actions = values.find(row => row[0] === 'operations')?.[1].split(',').map(item => item.trim() as ActionType) || [];
    configMap.autoAssign = (values.find(row => (row[0] || '').trim() === 'autoAssign')?.[1] || '')
        .split('/')
        .map(item => item.trim() as ActionType)
        .filter(Boolean);
    log.doLog(`getConfigMapping: autoAssign ${configMap.autoAssign.join('/') || '(none)'}`);
    if (!configMap.groupName) { 
        log.doLog('getConfigMapping: Warning groupName not found in config');
    }
    if (!configMap.actions || configMap.actions.length === 0) {
        log.doLog('getConfigMapping: Warning actions not found in config');
     }
    values.forEach(row => {
        if (row[0] === 'mapping') {
            configMap.taskLongToShortNameMapping[row[1]] = {
                shortNameOnFreedCamp: row[2].trim() as ActionType,
                copyFileFrom: row[3]?.trim(),
                copyFileTo: row[4]?.trim(),
                replaceInTemplateCopiedFile: row[5]?.trim(),
                updateCopiedFileToColumnName: row[6]?.trim()
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
        if (projectInfo && projectInfo.shortNameOnFreedCamp) {
            mapping[projectInfo.shortNameOnFreedCamp] = {
                project_id: proj.project_id,
                priority: 2,                
            }

            groupAndMainProjectMapping.shortProjectNameToProjectId[projectInfo.shortNameOnFreedCamp] = {
                project_id: proj.project_id,
                mappingConfig: projectInfo,
                // ========== COMMENTED OUT: isTaskEnabledForEnglish check - now using hasEnglishTemplate from UI checkbox instead ==========
                // isTaskEnabledForEnglish: projectInfo.isTaskEnabledForEnglishFromSheetConfig !== 'N',
                // ========== END COMMENTED OUT ==========
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
        if (projectInfo.shortNameOnFreedCamp && !mapping[projectInfo.shortNameOnFreedCamp]) {
            missingKeys.push(`${longName} (shortNameOnFreedCamp: ${projectInfo.shortNameOnFreedCamp})`);
        }
    }
    
    if (missingKeys.length > 0) {
        throw new Error(`Missing project mappings for the following keys: ${missingKeys.join(', ')}`);
    }
    
    return mapping;
}

export interface FreeCampAndUpdateOperations {
    getFreedCampToken: (loginParams: FreedCampLoginParams, forceReload: boolean) => Promise<LoginResponse>;
    getFreedCampProcessor: (loginToken: LoginResponse) => FreedCampProcessor;
}

export function getFreeCampAndUpdateOperations(freedCampOps: FreedCampOps): FreeCampAndUpdateOperations {
    async function getFreedCampToken(loginParams: FreedCampLoginParams) { //forceReload: boolean not used
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
    userNameToInfoMap: { [shortNameOnFreedCamp: string]: IUserInfo };
    userIdToInfoMap: { [userId: string]: IUserInfo };
    projectGroupMapping: IActionToProjectIdMapping;
    loginToken: LoginResponse;
    freedCampTasksByAction: { [key in ActionType]: IProjectTasksResult };
}
export async function combineOpsConfigWithFreedCampData(opsConfig: IOpsConfig, freedCampOps: FreeCampAndUpdateOperations, freedCampLoginParams: FreedCampLoginParams, log: DebugLog, forceReload: boolean)
:Promise<ICombinedOpsAndFreeCampData>
{
    const loginToken = await freedCampOps.getFreedCampToken(freedCampLoginParams, forceReload);
    const pr = freedCampOps.getFreedCampProcessor(loginToken);
    log.doLog('processOperation: got processor with login');
    
    const userData = await pr.getSessionCurrentData();
    const userIdToInfoMap: { [userId: string]: IUserInfo } = {};
    if (!userData.data.users) {
        log.doLog('Error: getSessionCurrentData did not return user data: ' + (userData as any).msg);
        throw new Error('getSessionCurrentData did not return user data: ' + (userData as any).msg);
    }
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


/** SharePoint subfolders under each article's mainFolder. Names must match the library exactly. */
export const SHAREPOINT_ARTICLE_FOLDERS = {
    original: '0. Original Materials',
    firstProof: '1 Text Editor & First Proofread',
    secondProof: '2 Second Proofread',
    artEditor: '3 Art Editor',
    publish: '4 Publish',
    media: '5 Media',
} as const;

/** A mapping row's copy-from value is a main-sheet column (old style) or a folder under mainFolder. */
export function isCopySourceFolder(opsConfig: IOpsConfig, copyFileFrom: string): boolean {
    return !opsConfig.headers.includes(copyFileFrom.trim());
}

export function rowHasDoctrineReview(
    operation: IOperationWithLineNumberAndParentTaskId,
    creatingDoctrineReview?: boolean,
): boolean {
    if (creatingDoctrineReview) return true;
    const taskId = String(operation[getTaskIdColumnName('教审' as ActionType)] || '').trim();
    if (taskId && taskId !== 'TaskIsEnglishAndIsDisabledForEnglish') return true;
    return !!(operation as IOperationWithLineNumberAndParentTaskId & { [key: string]: unknown })['教审 FreeCamp Item'];
}

export type ActionFilePlan = {
    sourceFolders?: string[];
    fallbackFolders?: string[];
    sourceColumn?: string;
    destFolder?: string;
    /** When false, the picked file is not used as {article}. Use linkPlaceholder instead. */
    usePickedAsArticle?: boolean;
    linkPlaceholder?: string;
    fileExtensions?: string[];
};

/** Where the Create dialog lists files, and whether creating the task should copy one. */
export function getActionFilePlan(
    action: ActionType,
    _operation: IOperationWithLineNumberAndParentTaskId,
    opsConfig: IOpsConfig,
    _creatingDoctrineReview?: boolean,
): ActionFilePlan | undefined {
    if ((action as string) === '教审') {
        return { sourceFolders: [SHAREPOINT_ARTICLE_FOLDERS.original] };
    }
    if ((action as string) === '校对') {
        return {
            sourceFolders: [SHAREPOINT_ARTICLE_FOLDERS.original, SHAREPOINT_ARTICLE_FOLDERS.firstProof],
            destFolder: SHAREPOINT_ARTICLE_FOLDERS.firstProof,
        };
    }
    if ((action as string) === '二校') {
        return {
            sourceFolders: [SHAREPOINT_ARTICLE_FOLDERS.firstProof, SHAREPOINT_ARTICLE_FOLDERS.secondProof],
            destFolder: SHAREPOINT_ARTICLE_FOLDERS.secondProof,
        };
    }
    if ((action as string) === '美编') {
        return {
            sourceFolders: [SHAREPOINT_ARTICLE_FOLDERS.secondProof],
            fallbackFolders: [SHAREPOINT_ARTICLE_FOLDERS.firstProof],
            destFolder: SHAREPOINT_ARTICLE_FOLDERS.artEditor,
        };
    }
    if ((action as string) === '语审') {
        return {
            sourceFolders: [SHAREPOINT_ARTICLE_FOLDERS.media],
            usePickedAsArticle: false,
            linkPlaceholder: 'voice_file',
            fileExtensions: ['.mp3'],
        };
    }
    const mapping = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action]?.mappingConfig;
    const name = mapping?.copyFileFrom?.trim();
    if (!name) return undefined;
    if (isCopySourceFolder(opsConfig, name)) {
        return { sourceFolders: [name], destFolder: mapping.copyFileTo?.trim() || undefined };
    }
    if (mapping.copyFileTo) {
        return { sourceColumn: name, destFolder: mapping.copyFileTo.trim() };
    }
    return undefined;
}

function isChineseTitle(title: string): boolean {
    return /[\u4e00-\u9fff]/.test(title);
}

/** English titles go in front ("Brother John"), Chinese titles after ("陈华弟兄"). */
function authorWithTitle(author: string | undefined, titleOfAddress: string | undefined): string {
    const name = (author || '').trim();
    const title = (titleOfAddress || '').trim();
    if (!name || !title) return name;
    if (isChineseTitle(title)) {
        return name.endsWith(title) ? name : `${name}${title}`;
    }
    return name.toLowerCase().startsWith(title.toLowerCase()) ? name : `${title} ${name}`;
}

function getOperationInfo(
    operation: IOperationWithLineNumberAndParentTaskId,
) {
    const article = operation['文章名'];
    //const fileName = operation['文件'];
    const infos: OperationInfo = {
        author: authorWithTitle(operation['作者'], operation['称呼']),
        slug: operation['slug'],
            article,
            //link: operation['文章链接'],
            email: operation['作者电邮'],
        category: operation['文章类别'],
            mainFolder: operation['mainFolder'],
            //校对: operation['校对'],
            editor: '',            
            //isEnglishOnly: !/[\u4e00-\u9fff]/.test(fileName)?'Y':'N', // Check if article contains Chinese characters
        };

        // Check if article is English-only (no Chinese characters)
    //const groupAndMainProjectMapping = combined.opsConfig.groupAndMainProjectMapping;
        // ========== COMMENTED OUT: isTaskEnabledForEnglish check - now using hasEnglishTemplate from UI checkbox instead ==========
        // Old logic: skip action if disabled for English and editor is English-only (Brother/Sister)
        // New logic: show all actions, let UI checkbox determine if English template should be used
        /*
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
        */
        // ========== END COMMENTED OUT ==========
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
    // const actionToTitleToProjectAndTaskIdMap: {
    //   [action in ActionType]?: {
    //     [title: string]: {
    //       projectId: string;
    //       taskId: string;
    //       task: IOperationWithLineNumberAndParentTaskId;
    //     };
    //   };
    // } = {};
    console.log('Checking for done sub-tasks to mark main tasks as done...');
    //   opsConfig.groupAndMainProjectMapping.actions.forEach(action => {
    //     const actionInfo = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[action];
    //     if (actionInfo && actionInfo.subTaskOf) {
    //       const subTaskOfAction = actionInfo.subTaskOf;
    //       projectList.forEach(item => {
    //         if (item[getTaskIdColumnName(subTaskOfAction)] === 'done' && !item[getTaskIdColumnName(action)]) {
    //           console.log(`EARNNN ${item.文件} ${item.文章名} ${action} as done for line ${item.itemPositionOnSheet} because ${subTaskOfAction} is done`);
    //           const prj = opsConfig.groupAndMainProjectMapping.shortProjectNameToProjectId[subTaskOfAction];
    //           if (prj) {
    //             console.log(`Debugrm parent project ${subTaskOfAction} project id ${prj.project_id} to check if all sub-tasks are done`);
    //             if (!actionToTitleToProjectAndTaskIdMap[subTaskOfAction]) actionToTitleToProjectAndTaskIdMap[subTaskOfAction] = {};                
    //             actionToTitleToProjectAndTaskIdMap[subTaskOfAction]![item.文件] = {
    //               projectId: prj.project_id,
    //               taskId: '',
    //               task: item,
    //             };
    //           }
    //         }
    //       });
    //     }
    //   });
    
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
                        let warnSTr = itm.文件;
                        if (!warnSTr) {
                            warnSTr = 'itm.文件 is empty';
                        }
                        // Known leftover: 发布 titles (二) and (三) both contain this 文件, and the tasks are no longer findable in FreedCamp.
                        const knownQuietDuplicate = (itm.文件 || '').includes('行走在圣经的字里行间');
                        if (!knownQuietDuplicate) {
                            log.doLog(`Warning: multiple FreedCamp tasks found for '${warnSTr}' in action ${action}, assigning the first one. Task IDs: ${freedCampTsks.map(t => t.id).join(', ')}`, true);
                        }
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
        getOperationInfo(item);
        const syncFreeCompToSheetParts = buildSyncUpdates(item, combinedOpsAndData);
        item.syncFreeCampToSheetData = syncFreeCompToSheetParts;
        const publishItem = item[`发布 FreeCamp Item`];
        let publishCompleted = false;
        if (publishItem) {
            publishCompleted = publishItem.completed_ts !== null && publishItem.completed_ts !== undefined && publishItem.status_title === 'Completed';
        }
        const taskDone = item['发布'] === 'done';
        item.isFinished = item.done === 'Y';//(publishCompleted || taskDone) && !item.syncFreeCampToSheetData;
        item.noNeedToCreate = taskDone;
        item.publishCompleted = publishCompleted;
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
    //here editor is shortNameOnFreedCamp
    const editorLookup = combined.opsConfig.editorInfoAry.find(e=>e.shortNameOnFreedCamp === editor);
    if (!editorLookup) {
        return '';
    }

    const prettyName = editorLookup.fullNamePN || editorLookup.shortNameOnFreedCamp;
    const isEnglishOnly = isEditorEnglish(editorLookup);
    return isEnglishOnly
        ? `${editorLookup.title} ${prettyName}`
        : `${prettyName}${editorLookup.title}`;
}

/** `{voice}` in a template becomes 男 or 女 from 称呼; other titles leave the choice to the editor. */
function voiceForTitleOfAddress(titleOfAddress: string | undefined): string {
    const title = (titleOfAddress || '').trim().toLowerCase();
    if (title === '弟兄' || title === 'brother') return '男';
    if (title === '姐妹' || title === 'sister') return '女';
    return '男/女';
}

function personKey(value: string): string {
    return value.normalize('NFKC').replace(/[.\s]/g, '').toLowerCase();
}

function findFreedCampUserForEditor(editor: string, combined: ICombinedOpsAndFreeCampData): IUserInfo | undefined {
    const direct = combined.userNameToInfoMap[editor];
    if (direct) {
        return direct;
    }
    const editorLookup = combined.opsConfig.editorInfoAry.find(info => info.shortNameOnFreedCamp === editor || info.fullNamePN === editor);
    const candidates = [editor];
    if (editorLookup?.fullNamePN) {
        candidates.push(editorLookup.fullNamePN);
    }
    if (editorLookup?.shortNameOnFreedCamp) {
        candidates.push(editorLookup.shortNameOnFreedCamp);
        const displayName = getEditorNameForAction(editorLookup.shortNameOnFreedCamp, combined);
        if (displayName) {
            candidates.push(displayName);
        }
    }
    const keys = new Set(candidates.map(personKey).filter(Boolean));
    return Object.values(combined.userNameToInfoMap).find(user => keys.has(personKey(user.full_name)));
}

export interface IEditorWorkload {
    openCount: number;
    completedCount: number;
    lastAssignedTs: number;
    lastCompletedTs: number;
    matchedUser: boolean;
}

export function getEditorWorkload(editor: string, action: ActionType, combined: ICombinedOpsAndFreeCampData): IEditorWorkload {
    const workload: IEditorWorkload = { openCount: 0, completedCount: 0, lastAssignedTs: 0, lastCompletedTs: 0, matchedUser: false };
    const user = findFreedCampUserForEditor(editor, combined);
    if (!user) return workload;
    workload.matchedUser = true;
    const DEFAULT_DUE_DAYS_SECONDS = 7 * 24 * 60 * 60;
    for (const task of combined.freedCampTasksByAction[action]?.tasks || []) {
        if (task.assigned_to_id !== user.user_id) continue;
        const assignedTs = Number(task.created_ts) || Number(task.start_ts)
            || (Number(task.due_ts) ? Number(task.due_ts) - DEFAULT_DUE_DAYS_SECONDS : 0);
        workload.lastAssignedTs = Math.max(workload.lastAssignedTs, assignedTs);
        if (task.status_title === 'Completed') {
            workload.completedCount++;
            workload.lastCompletedTs = Math.max(workload.lastCompletedTs, Number(task.completed_ts) || 0);
        } else {
            workload.openCount++;
        }
    }
    return workload;
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


export interface ActionsToPerformInfo {
    action: ActionType;
    editor: string;
    editorName: string;
    template: string;
    hasEnglishTemplate: boolean;
    hasAITemplate: boolean;
}
/**
 * Returns the list of actions that still need to be performed for a given operation.
 * Mirrors the filtering logic in processOperation:
 *   - Skips actions whose task ID is already set.
 *   - Skips actions disabled for English-only articles.
 *   - When an action has an AI template available, the create dialog exposes an AI checkbox.
 *     If selected, processOperation will use the AI template instead of the regular template.
 */
export function getActionsToPerform(
    operation: IOperationWithLineNumberAndParentTaskId,
    combined: ICombinedOpsAndFreeCampData,
    fileName: string, //fileName, for log only
    log: DebugLog,
): ActionsToPerformInfo[] {
    const { groupAndMainProjectMapping, templates } = combined.opsConfig;

    const excludedActions = new Set<ActionType>();
    for (const action of groupAndMainProjectMapping.actions) {
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
        // ========== COMMENTED OUT: isTaskEnabledForEnglish check - now using hasEnglishTemplate from UI checkbox instead ==========
        // Old logic: skip action if disabled for English and editor is English-only (Brother/Sister)
        // New logic: show all actions, let UI checkbox determine if English template should be used
        /*
        const editor = operation[action];
        const editorLookup = combined.opsConfig.editorInfoMap[editor];
        const isEnglishOnly = isEditorEnglish(editorLookup);
        if (actionConfig.isTaskEnabledForEnglish === false && isEnglishOnly) {
            log.doLog(`processOperation: skipping action ${action} for file ${fileName} as it is disabled from sheet config for English-only article`);
            operation[getTaskIdColumnName(action)] = 'TaskIsEnglishAndIsDisabledForEnglish';
            continue;
        }
        */
        // ========== END COMMENTED OUT ==========
        const editor = operation[action];
        const editorName = getEditorNameForAction(editor, combined) || 'EDITOR NOTSET';

        const existingTaskId = getExistingTaskId(action, operation);
        if (existingTaskId) {
            log.doLog(`processOperation: skipping action ${action} for file ${fileName} as task ID ${existingTaskId} exists`);
            taskIds.push(existingTaskId);
            continue;
        }
        const regularAction = templates[action];
        const AIActionName = `${action} AI` as ActionType;
        const aiAction = templates[AIActionName];
        let template1 = regularAction.template;
        const hasEnglishTemplate = !!regularAction.templateEnglish || !!aiAction?.templateEnglish;
        const hasAITemplate = !!aiAction;

        actionsToPerform.push({
            action,
            editorName,
            editor,
            template: template1,
            hasEnglishTemplate,
            hasAITemplate,
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
    debug_Prefix: string = '',
    /** When set, only FreedCamp tasks for these action keys are created (must be a subset of `getActionsToPerform`). */
    selectedActions?: ReadonlySet<ActionType>,
    /** Map of action -> whether to use English template (if available). */
    useAITemplateMap?: Map<ActionType, boolean>,
    useEnglishTemplateMap?: Map<ActionType, boolean>,
    msToken?: string,
    /** File picked in the Create dialog to copy for an action, in place of the copy-from column's link. */
    sourceFileOverride?: Map<ActionType, SharePointFileRef>,
): Promise<string[]> {
    const { groupAndMainProjectMapping } = combined.opsConfig;    
    log.doLog('processOperation: got processor with login');
    
    const taskIds = [];
    
    const pr = freedCampOps.getFreedCampProcessor(combined.loginToken);
    {
        const fileName = operation['文件'];
        const infos: OperationInfo = getOperationInfo(operation);

        // Check if article is English-only (no Chinese characters)
        
        const actionsToPerform = getActionsToPerform(operation, combined, fileName, log);
        const actionsToRun =
            selectedActions === undefined
                ? actionsToPerform
                : actionsToPerform.filter((a) => selectedActions.has(a.action));
        //const projectGroupMapping = getProjectGroupMapping();        
        //for (const action of combined.opsConfig.groupAndMainProjectMapping.actions) {
        for (const actionP of actionsToRun) {
            const action = actionP.action;
            const actionConfig = groupAndMainProjectMapping.shortProjectNameToProjectId[action];           
            infos.editor = actionP.editorName;
            // ========== BEFORE ENGLISH REFACTOR: old code that auto-selected template based on editor type ==========
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
            // ========== END BEFORE ENGLISH REFACTOR ==========
            let template1 = actionP.template;
            const regularAction = combined.opsConfig.templates[action];
            const AIActionName = `${action} AI` as ActionType;
            const aiAction = combined.opsConfig.templates[AIActionName];
            const useAIAction = !!useAITemplateMap?.get(action);
            const curTemplateActionInfo = useAIAction && aiAction ? aiAction : regularAction;
            if (useAIAction && !aiAction) {
                log.doLog(`processOperation: AI template requested for action ${action}, but ${AIActionName} was not found. Falling back to regular template.`, true);
            }
            if (useEnglishTemplateMap?.get(action) && curTemplateActionInfo.templateEnglish) {
                template1 = curTemplateActionInfo.templateEnglish;
            } else {
                template1 = curTemplateActionInfo.template;
            }
            const editor = operation[action];
            //const link = action === '校对' ? infos.link : undefined;            
            
            // if oneDriveDirInfos exists, find all matches in template that matches  {media_link:XXXX}, and if XXXX matches any of the oneDriveDirInfos.name, replace {media_link:XXXX} 
            // with the corresponding oneDriveDirInfos.webUrl
            if (oneDriveDirInfos && oneDriveDirInfos.length > 0) {
                template1 = template1.replace(/\{media_link:([^}]+)\}/g, (match, name) => {
                    const found = oneDriveDirInfos.find(info => info.name === name);
                    return found ? found.webUrl : match;
                });
            }            

            const mappingConfig = actionConfig.mappingConfig;
            const filePlan = getActionFilePlan(action, operation, combined.opsConfig, selectedActions?.has('教审' as ActionType));
            const pickedSource = sourceFileOverride?.get(action);
            const sourceValue = filePlan?.sourceColumn
                ? ((operation as any)[filePlan.sourceColumn] as string | undefined)?.trim()
                : '';
            if ((filePlan?.sourceFolders?.length || filePlan?.sourceColumn) && !pickedSource && !sourceValue) {
                log.doLog(
                    filePlan?.linkPlaceholder
                        ? `processOperation: ${action} has no file picked. {${filePlan.linkPlaceholder}} will be left empty.`
                        : `processOperation: ${action} has no file picked.`,
                    true,
                );
            }
            let copyLink = sourceValue
                || (mappingConfig?.replaceInTemplateCopiedFile ? operation['文章链接'] : '');
            if (pickedSource?.webUrl && !filePlan?.destFolder && filePlan?.usePickedAsArticle !== false) {
                copyLink = pickedSource.webUrl;
            }
            const destFolder = filePlan?.destFolder || mappingConfig?.copyFileTo;
            if (destFolder && (pickedSource || sourceValue)) {
                const destValue = operation.mainFolder
                    ? await joinSharePointPath(msToken!, operation.mainFolder, destFolder)
                    : undefined;
                if (destValue?.path?.trim()) {
                    if (!msToken) {
                        throw new Error('MS Graph token is required to copy SharePoint files for action ' + action);
                    }
                    if (pickedSource) {
                        log.doLog(`processOperation: ${action} copies ${pickedSource.name} to ${destFolder}`);
                    }
                    const copiedFileWebUrl = await copySharePointFile(msToken, pickedSource ?? sourceValue!, destValue.path.trim(), log);
                    if (copiedFileWebUrl) {
                        if (mappingConfig?.updateCopiedFileToColumnName) {
                            operation[mappingConfig.updateCopiedFileToColumnName as ActionType] = copiedFileWebUrl;
                            await updateAnyColumn(ops, combined.opsConfig, operation, mappingConfig.updateCopiedFileToColumnName, copiedFileWebUrl, log);
                        }
                        copyLink = copiedFileWebUrl;
                    }
                }
            }

            const titleOfAddress = (operation['称呼'] || '').trim();
            if (titleOfAddress) {
                template1 = isChineseTitle(titleOfAddress)
                    ? template1.replaceAll(`{author}${titleOfAddress}`, '{author}')
                    : template1.replaceAll(`${titleOfAddress} {author}`, '{author}');
            }
            //the above do action mapping article replacements, 
            for (const replaceItem of ['editor', 'author', 'email', 'article', 'category', 'slug']) {
                //!link because if it is link it is done below with link
                if (replaceItem === 'article') {
                    if (copyLink) {
                        template1 = template1.replaceAll('{article}', `<a href="${copyLink}">${infos['article']}</a>`);
                    } else {
                        template1 = template1.replaceAll('{article}', `${infos.article}`);
                    }
                    continue;
                }
                //console.log('debugreplace!!!!!!!!!!!', replaceItem, infos[replaceItem as keyof OperationInfo]);
                template1 = template1.replaceAll(`{${replaceItem}}`, infos[replaceItem as keyof OperationInfo]);
            }
            template1 = template1.replaceAll('{voice}', voiceForTitleOfAddress(operation['称呼']));
            const needsRelated = template1.includes('{相关文章}');
            const needsAbstract = template1.includes('{文章摘要}');
            if (needsRelated || needsAbstract) {
                let searchTerms: string[] = [];
                let relatedLanguage = detectArticleLanguage(infos.article, fileName);
                const articleMeta = {
                    title: infos.article,
                    author: infos.author,
                    category: infos.category,
                };
                const {
                    readPublishArticleText,
                    extractArticleSearchTerms,
                    extractArticleAbstract,
                    formatArticleAbstractHtml,
                } = await import('../src/lib/readPublishArticle');
                if (msToken && operation.mainFolder) {
                    try {
                        const article = await readPublishArticleText(msToken, operation.mainFolder, log);
                        relatedLanguage = detectArticleLanguage(article.text);
                        searchTerms = extractArticleSearchTerms(article.text, articleMeta);
                        log.doLog(`processOperation: related-article search from ${article.source} (${article.text.length} chars, ${relatedLanguage}): ${searchTerms.join(' | ')}`);
                        if (needsAbstract) {
                            const { abstract, foundParagraph } = extractArticleAbstract(article.text, articleMeta);
                            template1 = template1.replaceAll('{文章摘要}', formatArticleAbstractHtml(abstract));
                            log.doLog(`processOperation: {文章摘要} ${foundParagraph ? 'picked from article' : 'written'} (${abstract.length} chars)`);
                        }
                    } catch (error) {
                        const reason = (error as Error)?.message || String(error);
                        log.doLog(`processOperation: could not read publish 文章: ${reason}`, true);
                        if (needsAbstract) {
                            template1 = template1.replaceAll('{文章摘要}', formatArticleAbstractHtml('没有找到授权的文章'));
                        }
                    }
                } else if (needsAbstract) {
                    template1 = template1.replaceAll('{文章摘要}', formatArticleAbstractHtml('没有找到授权的文章'));
                }
                if (needsRelated) {
                    try {
                        const related = searchTerms.length
                            ? await findRelatedEnyuArticles(searchTerms, infos.slug, 5, relatedLanguage)
                            : [];
                        log.doLog(`processOperation: related-article language ${relatedLanguage}`);
                        template1 = template1.replaceAll('{相关文章}', formatRelatedArticlesHtml(related));
                        log.doLog(`processOperation: {相关文章} filled with ${related.length} post(s) for ${action}`);
                    } catch (error) {
                        log.doLog(`processOperation: could not look up 相关文章: ${(error as Error)?.message || String(error)}`, true);
                        template1 = template1.replaceAll('{相关文章}', '（未能查找相关文章）');
                    }
                }
            }
            if (pickedSource?.webUrl && filePlan?.linkPlaceholder) {
                const voiceLink = `<a href="${pickedSource.webUrl}">${pickedSource.name}</a>`;
                template1 = template1.replaceAll(`{${filePlan.linkPlaceholder}}`, voiceLink);
                if (pickedSource.folderName) {
                    template1 = template1.replace(new RegExp(`\\{media_link:${pickedSource.folderName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\}`, 'g'), pickedSource.webUrl);
                }
            }
            //if (link) {
            //    template1 = template1.replaceAll('{article}', `<a href="${infos['link']}">${infos['article']}</a>`);
            //}

            let projectGrpoup = combined.projectGroupMapping[action];            
            let taskTitle = `${debug_Prefix}${fileName}`
            //const subTaskOf = actionConfig.subTaskOf;
            // if (subTaskOf) {
            //     let h_parent_id = getExistingTaskId(subTaskOf, operation);
            //     if (!h_parent_id || h_parent_id === 'done') {
            //         h_parent_id = operation[getParentTaskIdColumnName(subTaskOf)] as string;
            //         if (!h_parent_id) {
            //             log.doLog(`processOperation: Warning no parent task ID found for subtask action ${action} with parent action ${subTaskOf} for file ${fileName}`);
            //             throw new Error(`No parent task ID found for subtask action ${action} with parent action ${subTaskOf}`);
            //         }

            //     }
            //     projectGrpoup = { ...combined.projectGroupMapping[subTaskOf] };
            //     projectGrpoup.h_parent_id = h_parent_id;
            //     projectGrpoup.description = template1;
            //     taskTitle = action;
            // }
            const dueDateKey = `${action} Due Date` as DueDateKeys;
            let due_date = operation[dueDateKey];

            const today = new Date();
            const todayYyyyMmDd = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
            const start_date = todayYyyyMmDd;

            if (due_date) {                
                const dueDateObj = new Date(due_date);
                if (!isNaN(dueDateObj.getTime())) {
                    //postParams.due_ts = Math.floor(dueDateObj.getTime() / 1000);
                    due_date = `${dueDateObj.getFullYear()}-${String(dueDateObj.getMonth() + 1).padStart(2, '0')}-${String(dueDateObj.getDate()).padStart(2, '0')}`;
                }
            }
            const taskRes = await pr.createTask(projectGrpoup, taskTitle, start_date, due_date);
            const taskId = taskRes.id.toString();
            console.log(`Created task action ${action} ${taskId} for file ${fileName}`);
            log.doLog(`processOperation: created task action ${action} ${taskId} for file ${fileName}`);
            operation[getTaskIdColumnName(action)] = taskId;
            await taskIdUpdater(ops, combined.opsConfig, action, operation, log);
            taskIds.push(taskId);
            
            
            const postParams: ProjectTaskParams = {
                description: template1,
                "priority": 2,
                time_from: '00:00',
                time_to: '00:00'
            };
            
            if (groupAndMainProjectMapping.autoAssign.includes(action)) {
                if (!editor) {
                    log.doLog(`processOperation: autoAssign ${action} has no editor on the row. Task left unassigned.`, true);
                } else {
                    const userInfo = findFreedCampUserForEditor(editor, combined);
                    if (userInfo) {
                        postParams.assigned_to_id = userInfo.user_id;
                        log.doLog(`processOperation: autoAssign ${action} to ${userInfo.full_name}`);
                    } else {
                        log.doLog(`processOperation: autoAssign ${action} could not match a FreedCamp user for '${editor}'. Task left unassigned.`, true);
                    }
                }
            }            
            log.doLog(`processOperation: prepared post params for task ${taskId} action ${action} with ${JSON.stringify(postParams)}`);
            //due_date
            //if (!subTaskOf) {
            await pr.doPostAttachment(taskId, postParams);
            //}
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
    
    const loginToken = await freedCampOps.getFreedCampToken(freedCampLoginParams, false);
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
    const combined = await combineOpsConfigWithFreedCampData(opsAndTemplates, fops, freedCampLoginParams, log, false);
    const ids = await processOperation(ops, fops, combined, validOperation, [], log, prefix);
    return {
        ids, 
        ...opsAndTemplates,
    }
}
