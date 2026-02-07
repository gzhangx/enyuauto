import type { ActionType } from "./types";

import * as gs from '@gzhangx/googleapi';
export type DueDateKeys = `${ActionType} Due Date`;

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

export type OperationWithDueDates = Operation & {
    [K in DueDateKeys]: string;
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
        taskIdLine: number;
        taskIdUpdater: (newTaskId: string) => Promise<void>;
        existingTaskId: string;
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
            subTaskOf?: ActionType;
        }; 
    };
    shortProjectNameToProjectId: { //populated later after we login to freedcamp
        [key in ActionType]: {
            project_id: string;
            subTaskOf?: ActionType;
        }; //'校对': { "project_id": "3696514" }
    }
}

export interface IEditorInfoMap { [key: string]: IEditorInfo };
export interface OperationAndTemplates {
    validOperation: OperationWithDueDates;
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
    operationList: OperationWithDueDates[];
    groupAndMainProjectMapping: IGroupAndMainProjectLongToShortNameMapping;
    editorInfoMap: IEditorInfoMap;
    headers: string[];
}