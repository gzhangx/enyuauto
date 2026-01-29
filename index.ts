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
  const log: mainOps.DebugLog = { operations: [] };
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
      throw new Error('Please provide a valid line number as the "line" query parameter.');
    } 
    const lineNumber = parseInt(lineNumberStr);
        
    const res = params['doit']?await mainOps.debug_main(lineNumber, log, opStr):'No operation performed';
    
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
      statusCode: 500,
      body: JSON.stringify({ 
        message: 'Error',
        error: error instanceof Error ? error.message : String(error),
        log: log.operations,
      }),
    };
  }
};
