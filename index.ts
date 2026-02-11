import * as mainOps from './web/shared/main_ops';
import * as util from './lib/util';
import { ProjectTaskParams, FreedCampProcessor, LoginResponse } from './web/shared/freedcampTypes';
import * as secs from './enyu_secs.json';
//aws logs tail /aws/lambda/enyu_auto --follow --region us-east-1
interface LambdaEvent {
  queryStringParameters: {
    action: string;
    subAction?: string;
    params: { [key: string]: string };
    test: string;
    line: string;
  };
  body: string;
  [key: string]: any;
}

interface LambdaResponse {
  statusCode: number;
  headers?: { [key: string]: string };
  body: string;
}
const headers = {
  'Content-Type': 'application/json'
}

export const handler = async (event: LambdaEvent): Promise<LambdaResponse> => {
  const operations: string[] = [];
  const log: mainOps.DebugLog = {     
    getOperations: () => operations,
    doLog: (msg: string) => {     
      operations.push(msg);
      console.log(msg);
    }
  };
  try {
    const body = event.body ? JSON.parse(event.body) : {};
    if (body.action === 'freedcamp') {
      return doFreedcampAction(body, log);
    }
    const ops = await mainOps.getSheetOps(secs.gsAuth);
    const params = event.queryStringParameters || {};
    
    const action = params['action'] || 'main' as 'main' | 'getList' | 'freedcamp';

    
    if (action === 'getList') {
      const list = await mainOps.getOpsAndMainList(ops, log);
      return {
        statusCode: 200,
        headers: headers,          
        body: JSON.stringify(list),
      };
    }
    let opStr: string | undefined;
    if (params['test']) {
      opStr = 'debug_remove_test_';
    }

    const lineNumberStr = params['line'];
    if (!lineNumberStr || !lineNumberStr.match(/^\d+$/)) {
      return {
        statusCode: 400,
        headers: headers,
        body: JSON.stringify(
          { error: 'Please provide a valid line number as the "line" query parameter.' }),
      };
    } 
    const lineNumber = parseInt(lineNumberStr);
        
    let res = 'No operation performed';
    if (action === 'main' || action === 'del') {
      try {
        if (action === 'del') opStr = 'del';
        console.log(`Performing operation: ${opStr} on line ${lineNumber}`);
        const fops = mainOps.getFreeCampAndUpdateOperations(util.freedCampOps);
        const mainCfg = await mainOps.getOpsAndMainList(ops, log);
        let sres: string[] = [];
        const operation = mainOps.getOperationFromLineNumber(mainCfg.operationList, lineNumber);
        if (!operation) {
          throw new Error(`No operation found for line number ${lineNumber}`);
        }
        for (const action of mainCfg.groupAndMainProjectMapping.actions) {
          const rr = await mainOps.deleteItemActionTask(ops, fops, mainCfg, operation, action, log);
          sres.push(rr || '');
        }
        res = sres.join('\n');
      } catch (error) {
        console.error('Operation error:', error);
        const errorMessage = error instanceof Error ? error.message : String(error);
        const errorStack = error instanceof Error ? error.stack : undefined;
        res = `Operation error: ${errorMessage}`;
        log.doLog(`ERROR: ${errorMessage}`);
        if (errorStack) {
          log.doLog(`STACK: ${errorStack}`);
        }
      }
    }
    
    const response: LambdaResponse = {
      statusCode: 200,
      headers: headers,
      body: JSON.stringify({
        message: 'Success',
        operation: opStr || 'main',
        res,
        params: event.queryStringParameters || {},
        log: operations,
      }),
    };
    return response;
  } catch (error) {
    console.log('General error in main:', error);
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ 
        message: 'Error',
        error: error instanceof Error ? error.message : String(error),
        log: operations,
      }),
    };
  }
};

// action: freedcamp
// subAction: login, getSessionCurrentData, createTask, deleteTask, doPostAttachment
// login: username, password
// createTask: projectId, title, description?, assignedToId?, dueDate?, parentId?
// deleteTask: taskId
// doPostAttachment: taskId, description?, assignedToId?, dueDate?
async function doFreedcampAction(params: { [key: string]: string; }, log: mainOps.DebugLog): Promise<LambdaResponse> {
  try {
    if (!params) throw new Error('Missing parameters for freedcamp action');
    log.doLog('Processing freedcamp action');
    
    let result: any = {};
    const subAction = params['subAction'];
    
    // For most actions, we need cookies (either from params or via login)
    let processor: FreedCampProcessor | undefined;
    if (subAction !== 'login') {
      let cookies: LoginResponse = {} as LoginResponse;
      
      // Check if cookies are provided in parameters
      if (params['cookies']) {
        cookies.Cookie = params['cookies'];
        log.doLog(`Using ${cookies.Cookie} cookies from parameters for action: ${subAction}`);
      } else {
        // If no cookies provided, try to login
        const username = params['username'] || '';
        const password = params['password'] || '';
        if (!username || !password) {
          return {
            statusCode: 400,
            headers: headers,
            body: JSON.stringify({
              message: 'Missing authentication',
              error: 'Either cookies parameter or username/password are required for this action',
            }),
          };
        }
        cookies = await util.freedCampOps.login({ username, password });
        log.doLog(`Logged in successfully for action: ${subAction}`);
      }
      
      processor = util.freedCampOps.getProcessor(cookies);
    }
    
    switch (subAction) {
      case 'login':
        const loginRes = await util.freedCampOps.login({
          username: params['username'] || '',
          password: params['password'] || '',
        });
        result = loginRes;
        break;
        
      case 'getSessionCurrentData':
        if (!processor) throw new Error('Processor not initialized');
        const sessionData = await processor.getSessionCurrentData();
        result = sessionData;
        log.doLog('Retrieved session current data');
        break;
        
      case 'createTask':
        if (!processor) throw new Error('Processor not initialized');
        const projectId = params['projectId'];
        const title = params['title'];
        if (!projectId || !title) {
          return {
            statusCode: 400,
            headers: headers,
            body: JSON.stringify({
              message: 'Missing required parameters',
              error: 'projectId and title are required for createTask',
            }),
          };
        }
        const taskParams: ProjectTaskParams = {
          project_id: projectId,
          priority: 2,
        };
        if (params['description']) taskParams.description = params['description'];
        if (params['assignedToId']) taskParams.assigned_to_id = params['assignedToId'];
        if (params['dueDate']) taskParams.due_date = params['dueDate'];
        if (params['parentId']) taskParams.h_parent_id = params['parentId'];
        
        const createRes = await processor.createTask(taskParams, title);
        result = createRes;
        log.doLog(`Created task with ID: ${createRes.id}`);
        break;
        
      case 'deleteTask':
        if (!processor) throw new Error('Processor not initialized');
        const taskId = params['taskId'];
        if (!taskId) {
          return {
            statusCode: 400,
            headers: headers,
            body: JSON.stringify({
              message: 'Missing required parameters',
              error: 'taskId is required for deleteTask',
            }),
          };
        }
        const deleteRes = await processor.deleteTask(taskId);
        result = deleteRes;
        log.doLog(`Deleted task with ID: ${taskId}`);
        break;
        
      case 'doPostAttachment':
        if (!processor) throw new Error('Processor not initialized');
        const attachTaskId = params['taskId'];
        if (!attachTaskId) {
          return {
            statusCode: 400,
            headers: headers,
            body: JSON.stringify({
              message: 'Missing required parameters',
              error: 'taskId is required for doPostAttachment',
            }),
          };
        }
        const attachParams: ProjectTaskParams = {
          priority: 2,
        };
        if (params['description']) attachParams.description = params['description'];
        if (params['assignedToId']) attachParams.assigned_to_id = params['assignedToId'];
        if (params['dueDate']) attachParams.due_date = params['dueDate'];
        
        const attachRes = await processor.doPostAttachment(attachTaskId, attachParams);
        result = attachRes;
        log.doLog(`Updated task ${attachTaskId} with attachment/params`);
        break;
        
      default:
        return {
          statusCode: 400,
          headers: headers,
          body: JSON.stringify({
            message: 'Invalid subAction',
            error: `Unknown subAction: ${subAction}. Valid actions: login, getSessionCurrentData, createTask, deleteTask, doPostAttachment`,
          }),
        };
    }
    
    return {
      statusCode: 200,
      headers: headers,
      body: JSON.stringify(result),
    };
  } catch (error) {
    log.doLog(`Freedcamp action error: ${error instanceof Error ? error.message : String(error)}`);
    return {
      statusCode: 500,
      headers: headers,
      body: JSON.stringify({
        message: 'Freedcamp action failed',
        error: error instanceof Error ? error.message : String(error),
        log: log.getOperations(),
      }),
    };
  }
}


