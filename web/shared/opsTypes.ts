import type { ISheetInfoSimple } from "@gzhangx/googleapi/lib/googleApi";
import type { ActionType } from "./types";

import * as gs from '@gzhangx/googleapi';
import type { ProjectTaskParams } from "./freedcampTypes";
export type DueDateKeys = `${ActionType} Due Date`;
export type CompleteDateKeys = `${ActionType} Complete Date`;
export type TaskIdKeys = `${ActionType} TaskId`;
export type FreeCampItemKeys = `${ActionType} FreeCamp Item`;
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
};

export interface OperationInfo {
    author: string;
    article: string;
    link: string;
    email: string;
    category: string;
    editor: string;
}


export type Templates = {
    [K in ActionType]: {
        template: string;
        templateEnglish: string;
        taskIdPos: number;
        completeDatePos: number;
        //taskIdLine: number;
        //taskIdUpdater: (newTaskId: string, lineNumber: number) => Promise<void>;
        //getExistingTaskId: (operation: OperationWithDueDates)=>string;
    }
};

export interface IGroupAndMainProjectLongToShortNameMapping {
    freedcampInfo: {
        username: string;
        password: string;
    };
    groupName: string; //EnYu_2026
    actions: ActionType[];
    taskLongToShortNameMapping: {
        [key: string]: {
            shortName: ActionType; //文字校对 (Editorial and Translation team) : 校对
            subTaskOfFromSheetConfig?: ActionType;
            isTaskEnabledFromSheetConfig?: '' | 'N';
        }; 
    };
    shortProjectNameToProjectId: { //populated later after we login to freedcamp
        [key in ActionType]: {
            project_id: string;
            subTaskOf?: ActionType;
            isTaskEnabled?: boolean;
        }; //'校对': { "project_id": "3696514" }
    }
}

export interface IEditorInfoMap { [key: string]: IEditorInfo };
export interface OperationAndTemplates {
    validOperation: IOperationWithLineNumberAndParentTaskId;
    templates: Templates;
    ops: gs.gsAccount.IGetSheetOpsReturn;
    editorInfoMap: IEditorInfoMap;
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
}


export interface IEditorInfo {
    title: string;
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