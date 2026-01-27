interface LambdaEvent {
  [key: string]: any;
}

interface LambdaResponse {
  statusCode: number;
  body: string;
}

export const handler = async (event: LambdaEvent): Promise<LambdaResponse> => {
  // TODO implement
  const response: LambdaResponse = {
    statusCode: 200,
    body: JSON.stringify('Hello from Lambda!'),
  };
  return response;
};
