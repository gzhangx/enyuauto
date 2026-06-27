import * as mainOps from './web/shared/main_ops';
import * as util from './lib/util';
import { ProjectTaskParams, FreedCampProcessor, LoginResponse } from './web/shared/freedcampTypes';
//import * as secs from './enyu_secs.json';
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
    if (body.action === 'wordpress') {
      return doWordpressAction(body, log);
    }
    //action below are no longer used.
    //const ops = await mainOps.getSheetOps(secs.gsAuth);
    const params = event.queryStringParameters || {};
    
    const action = params['action'] || 'main' as 'main' | 'getList' | 'freedcamp';

    
    if (action === 'getList') {
      //const list = await mainOps.getOpsAndMainList(ops, log);
      return {
        statusCode: 200,
        headers: headers,          
        body: 'no longer supported', //JSON.stringify(list),
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
        //const fops = mainOps.getFreeCampAndUpdateOperations(util.freedCampOps);
        //const mainCfg = await mainOps.getOpsAndMainList(ops, log);
        let sres: string[] = [];
        //const operation = mainOps.getOperationFromLineNumber(mainCfg.operationList, lineNumber);
        //if (!operation) {
          throw new Error(`No operation found for line number ${lineNumber}`);
        //}        
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
// getTasksForProjects: projectId, pageNumber
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
    let processor: FreedCampProcessor = null as unknown as FreedCampProcessor;
    if (subAction !== 'login') {
      let cookies: LoginResponse = {} as LoginResponse;
      
      // Check if cookies are provided in parameters
       const username = params['username'] || '';
      const password = params['password'] || '';
      if (params['cookies'] && !username && !password) {
        cookies.Cookie = params['cookies'];
        log.doLog(`Using ${cookies.Cookie} cookies from parameters for action: ${subAction}`);
      } else {
        // If no cookies provided, try to login       
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
        log.doLog('Performing login action');
        const loginRes = await util.freedCampOps.login({
          username: params['username'] || '',
          password: params['password'] || '',
        });
        result = loginRes;
        log.doLog('Performed login action ' + JSON.stringify(loginRes));
        break;
        
      case 'getSessionCurrentData':
        if (!processor) throw new Error('Processor not initialized');
        const sessionData = await processor.getSessionCurrentData();
        result = sessionData;
        log.doLog('Retrieved session current data');
        break;
      case 'getTasksForProjects':
        result = await processor.getTasksForProjects(params['projectId'], parseInt(params['pageNumber'] || '1'));
        break;
        
      case 'createTask':
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
        if (params['dueDate']) taskParams.due_ts = params['dueDate'] as unknown as number;
        if (params['parentId']) taskParams.h_parent_id = params['parentId'];
        
        const createRes = await processor.createTask(taskParams, title);
        result = createRes;
        log.doLog(`Created task with ID: ${createRes.id}`);
        break;
        
      case 'deleteTask':
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
          ...params,
        };
        if (params['description']) attachParams.description = params['description'];
        if (params['assignedToId']) attachParams.assigned_to_id = params['assignedToId'];
        if (params['dueDate']) attachParams.due_ts = params['dueDate'] as unknown as number;
        
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
      }),
    };
  }
}

const WP_BASE_URL = 'https://enyu.acccn.org/wp-json/wp/v2';

// action: wordpress
// subAction: uploadMedia, createPost
// uploadMedia: wpToken, filename, mimeType, b64
// createPost: wpToken, title, content, status?
async function doWordpressAction(params: { [key: string]: string }, log: mainOps.DebugLog): Promise<LambdaResponse> {
  try {
    if (!params) throw new Error('Missing parameters for wordpress action');
    const subAction = params['subAction'];
    const wpToken = params['wpToken'];
    if (!wpToken) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ message: 'Missing wpToken' }),
      };
    }
    const authHeader = wpToken;

    let result: any = {};

    switch (subAction) {
      case 'uploadMedia': {
        const filename = params['filename'];
        const mimeType = params['mimeType'];
        const b64 = params['b64'];
        if (!filename || !mimeType || !b64) {
          return {
            statusCode: 400,
            headers,
            body: JSON.stringify({ message: 'Missing required parameters', error: 'filename, mimeType, and b64 are required for uploadMedia' }),
          };
        }
        const binaryBuffer = Buffer.from(b64, 'base64');
        const res = await fetch(`${WP_BASE_URL}/media`, {
          method: 'POST',
          headers: {
            Authorization: authHeader,
            'Content-Disposition': `attachment; filename="${filename}"`,
            'Content-Type': mimeType,
          },
          body: binaryBuffer,
        });
        result = await res.json();
        log.doLog(`Uploaded media: ${filename}, status: ${res.status}`);
        break;
      }

      case 'createPost': {
        const title = params['title'];
        const content = params['content'];
        if (!title || !content) {
          return {
            statusCode: 400,
            headers,
            body: JSON.stringify({ message: 'Missing required parameters', error: 'title and content are required for createPost' }),
          };
        }
        const res = await fetch(`${WP_BASE_URL}/posts`, {
          method: 'POST',
          headers: {
            Authorization: authHeader,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            title,
            content,
            status: params['status'] || 'draft',
          }),
        });
        result = await res.json();
        log.doLog(`Created post: ${title}, status: ${res.status}`);
        break;
      }

      default:
        return {
          statusCode: 400,
          headers,
          body: JSON.stringify({
            message: 'Invalid subAction',
            error: `Unknown subAction: ${subAction}. Valid actions: uploadMedia, createPost`,
          }),
        };
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(result),
    };
  } catch (error) {
    log.doLog(`WordPress action error: ${error instanceof Error ? error.message : String(error)}`);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        message: 'WordPress action failed',
        error: error instanceof Error ? error.message : String(error),
      }),
    };
  }
}
