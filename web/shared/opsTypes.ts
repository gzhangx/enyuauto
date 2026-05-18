import type { ISheetInfoSimple } from "@gzhangx/googleapi/lib/googleApi";
import type { ActionType } from "./types";

import type { ProjectTaskParams } from "./freedcampTypes";

/** Abstract sheet/workbook data-access layer — implemented by GoogleSheetDataOps or MsExcelDataOps. */
export interface ISheetDataOps {
    readData(sheetName: string): Promise<{ values: string[][] }>;
    /**
     * Write `values` to a single rectangular block.
     * @param row  1-indexed data row (0 = header row, 1 = first data row).
     * @param col  0-indexed column number.
     */
    autoUpdateValues(sheetName: string, values: string[][], position: { row: number; col: number }): Promise<void>;
}
export type DueDateKeys = `${ActionType} Due Date`;
export type CompleteDateKeys = `${ActionType} Complete Date`;
export type TaskIdKeys = `${ActionType} TaskId`;
export type FreeCampItemKeys = `${ActionType} FreeCamp Item`;
export type AITemplateActionKeys = `${ActionType} AI`;
export type SyncUpdateItem = {
    sheetCol: string;
    value: string;
};

export type ISyncFreeCampToSheetData = {
    parts: string[];
    updates: SyncUpdateItem[];
};

export interface Operation {
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
    mainFolder: string;
    slug: string;
    done: string | 'Y' | 'N' | '';
}

export function getTaskIdColumnName(action: ActionType): TaskIdKeys {
    return `${action} TaskId`;
}

export function getCompleteDateColumnName(action: ActionType): CompleteDateKeys {
    return `${action} Complete Date`;
}

export type OperationWithDueDates = Operation & {
    [K in DueDateKeys]: string;    
} & {
    [k in TaskIdKeys]: string;
}& {
    [k in CompleteDateKeys]: string;
};

export type IOperationWithLineNumber = OperationWithDueDates & { itemPositionOnSheet: number; };

export type ParentTaskIdKeys = `${ActionType} ParentTaskId`;
export function getParentTaskIdColumnName(action: ActionType): ParentTaskIdKeys {
    return `${action} ParentTaskId`;
}
// this is not on sheet, but used to store parent task id in memory in case parent task is done and we have to retrive parent task id
export type IOperationWithLineNumberAndParentTaskId = IOperationWithLineNumber & {
    [k in ParentTaskIdKeys]?: string;
} & {
    [k in FreeCampItemKeys]?: ProjectTaskParams;
} & {
    [k in AITemplateActionKeys]?: string;
} & {
    syncFreeCampToSheetData?: ISyncFreeCampToSheetData;
    noNeedToCreate: boolean;  //if we alreadu published then we should not need to create any more
    isFinished: boolean;
};

export interface OperationInfo {
    author: string;
    slug: string;
    article: string;
    link: string;
    email: string;
    category: string;
    mainFolder: string;
    editor: string;
}


export type Templates = {
    [K in ActionType | AITemplateActionKeys]: {
        template: string;
        templateEnglish: string;
        taskIdPos: number;
        completeDatePos: number;
        //taskIdLine: number;
        //taskIdUpdater: (newTaskId: string, lineNumber: number) => Promise<void>;
        //getExistingTaskId: (operation: OperationWithDueDates)=>string;
    }
};

export interface ProjectActionMappingConfig {
    shortName: ActionType; //文字校对 (Editorial and Translation team) : 校对
    copyFileFrom?: string;
    copyFileTo?: string;
    replaceInTemplateCopiedFile?: string;
}

export interface IGroupAndMainProjectLongToShortNameMapping {
    groupName: string; //EnYu_2026
    actions: ActionType[];
    taskLongToShortNameMapping: {
        [key: string]: ProjectActionMappingConfig;
    };
    shortProjectNameToProjectId: { //populated later after we login to freedcamp
        [key in ActionType]: {
            project_id: string;
            mappingConfig: ProjectActionMappingConfig;
            // ========== COMMENTED OUT: isTaskEnabledForEnglish check - now using hasEnglishTemplate from UI checkbox instead ==========
            // isTaskEnabledForEnglish?: boolean;
            // ========== END COMMENTED OUT ==========
        }; //'校对': { "project_id": "3696514" }
    };
    actionExcludes: {
        [key in ActionType]?: ActionType[]; //if actionExcludes['语音制作 AI'] = ['语音编辑'], it means if 语音制作 AI Exists, then we should exclude 语音编辑
    };
}

export interface IEditorInfoMap { [key: string]: IEditorInfo };
export interface OperationAndTemplates {
    validOperation: IOperationWithLineNumberAndParentTaskId;
    templates: Templates;
    ops: ISheetDataOps;
    editorInfoMap: IEditorInfoMap;
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
}


export interface IEditorInfo {
    title: string;   //if title is brother then it is english
    shortName: string;
    email: string;
    task: string;
    print_name: string; //the full chinese name for printing
}

export interface IOpsConfig {
    operationList: IOperationWithLineNumberAndParentTaskId[];
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
    editorInfoMap: IEditorInfoMap;
    headers: string[];
    templates: Templates;    
}



export interface ISheetInfoCache {
    getCachedSheetInfo(): ISheetInfoSimple[] | null;
    setCacheSheetInfo(data: ISheetInfoSimple[]): void;
}