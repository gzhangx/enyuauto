import * as mainOps from './lib/main_ops';

interface LambdaEvent {
  queryStringParameters?: {
    [key: string]: string;
  };
  [key: string]: any;
}

interface LambdaResponse {
  statusCode: number;
  body: string;
}

export const handler = async (event: LambdaEvent): Promise<LambdaResponse> => {
  try {
    const url = event.queryStringParameters?.url || '';
    
    let opStr: string | undefined;
    if (url.includes('test')) {
      opStr = 'debug_remove_test_';
    } else if (url.includes('del')) {
      opStr = 'del';
    }
    
    const res = await mainOps.main(opStr);
    
    const response: LambdaResponse = {
      statusCode: 200,
        body: JSON.stringify({
            message: 'Success',
            operation: opStr || 'main',
            res,
        }),
    };
    return response;
  } catch (error) {
    return {
      statusCode: 500,
      body: JSON.stringify({ 
        message: 'Error',
        error: error instanceof Error ? error.message : String(error)
      }),
    };
  }
};
