import * as mainOps from './lib/main_ops';

interface LambdaEvent {
  queryStringParameters: {
    [key: string]: string;
  };
  [key: string]: any;
}

interface LambdaResponse {
  statusCode: number;
  body: string;
}

export const handler = async (event: LambdaEvent): Promise<LambdaResponse> => {
  const log: mainOps.DebugLog = { 
    operations: [], 
    doLog: (msg: string) => log.operations.push(msg) 
  };
  try {
    const params = event.queryStringParameters || {};
    
    let opStr: string | undefined;
    if (params['test']) {
      opStr = 'debug_remove_test_';
    } else if (params['del']) {
      opStr = 'del';
    }

    const lineNumberStr = params['line'];
    if (!lineNumberStr || !lineNumberStr.match(/^\d+$/)) {
      return {
        statusCode: 200,
        body: JSON.stringify(
          { error: 'Please provide a valid line number as the "line" query parameter.' }),
      };
    } 
    const lineNumber = parseInt(lineNumberStr);
        
    let res = 'No operation performed';
    if (params['doit']) {
      try {
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
    return {
      statusCode: 200,
      body: JSON.stringify({ 
        message: 'Error',
        error: error instanceof Error ? error.message : String(error),
        log: log.operations,
      }),
    };
  }
};
