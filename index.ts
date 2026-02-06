import * as mainOps from './lib/main_ops';
import * as util from './lib/util';
//aws logs tail /aws/lambda/enyu_auto --follow --region us-east-1
interface LambdaEvent {
  queryStringParameters: {
    [key: string]: string;
  };
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
  const log: mainOps.DebugLog = { 
    operations: [], 
    doLog: (msg: string) => {
      log.operations.push(msg);
      console.log(msg);
    }
  };
  try {
    const params = event.queryStringParameters || {};
    
    const action = params['action'] || 'main' as 'main' | 'getList' | 'freedcamp';

    if (action === 'freedcamp') {
      return doFreedcampAction(params, log);
    }
    if (action === 'getList') {
      const list = await mainOps.getOpsAndMainList(log);
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
        const boolRes = await mainOps.debug_main(lineNumber, log, opStr);
        res = boolRes ? 'Operation succeeded' : 'Operation failed';
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
        log: log.operations,
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
        log: log.operations,
      }),
    };
  }
};

async function doFreedcampAction(params: { [key: string]: string; }, log: mainOps.DebugLog): Promise<LambdaResponse> {
  try {
    log.doLog('Processing freedcamp action');
    
    let result: any = {};
    switch (params['subAction']) {
      case 'login':
        const loginRes = await util.login({
          username: params['username'] || '',
          password: params['password'] || '',
        });
        result = { cookies: loginRes };
        break;
    }
    // Example implementation - adjust based on your actual freedcamp integration needs
    
    
    return {
      statusCode: 200,
      headers: headers,
      body: JSON.stringify({
        message: 'OK',
        result,        
      }),
    };
  } catch (error) {
    log.doLog(`Freedcamp action error: ${error instanceof Error ? error.message : String(error)}`);
    return {
      statusCode: 500,
      headers: headers,
      body: JSON.stringify({
        message: 'Freedcamp action failed',
        error: error instanceof Error ? error.message : String(error),
        log: log.operations,
      }),
    };
  }
}


