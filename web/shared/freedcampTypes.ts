export interface FreedCampLoginParams {
    username: string;
    password: string;
}

export interface LoginResponse {
    Cookie: string;
}

export interface ProjectTaskParams {
    id?: string;
    h_parent_id?: string;
    project_id?: string;
    //task_group_id: string;
    description?: string;
    assigned_to_id?: string;
    due_ts?: number;
    start_date?: string;
    due_date?: string;
    time_from?: '00:00',
    time_to?: '00:00',
    title?: string;
    status?: number;
    status_id?: number;
    status_title?: string; //'Completed': 1, No Progress: 0
    completed_ts?: number;
    assigned_to_fullname?: string;
    app_id?: string; //2
    "priority": 2;
    priority_title?: string; //'Medium'
}

export interface CreateTaskResponse {
    result: any;
    id: number;
}

export interface FreedCampProcessor {
    doPostAttachment: (taskId: string, params: ProjectTaskParams) => Promise<any>;
    createTask: (projectAndGroup: ProjectTaskParams, title: string) => Promise<CreateTaskResponse>;
    deleteTask: (taskId: string | number) => Promise<any>;
    getSessionCurrentData(): Promise<ICurrentSessionData>;
    getTasksForProjects(projectId: string, pageNumber: number): Promise<IProjectTasksResult>
}

export interface IUserInfo {
    user_id: string;
    full_name: string;
    email: string;
}

export interface IProjectInfo {
    project_id: string;
    //group_id: string;
    project_name: string; //文字二校 （Second Proofread)
    group_name: string; //EnYu_2026
}

export interface ICurrentSessionData {
    data: {
        users: IUserInfo[];
        projects: IProjectInfo[];
    }
}


export interface FreedCampOps {
    login: (prm: FreedCampLoginParams) => Promise<LoginResponse>;
    getProcessor: (loginResponse: LoginResponse) => FreedCampProcessor;
}

export interface IProjectTasksResult {
    tasks: ProjectTaskParams[];
}