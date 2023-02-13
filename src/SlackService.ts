export function callWebApi(token: string, apiMethod: string, payload: any) {
  const response = UrlFetchApp.fetch(`https://www.slack.com/api/${apiMethod}`, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {Authorization: `Bearer ${token}`},
    payload: payload,
  });
  Logger.log(`Web API (${apiMethod}) response: ${response}`);
  return response;
}

export function uploadFileToSlack(token: string, payload: any) {
  const endpoint = 'https://www.slack.com/api/files.upload';
  if (payload['file'] !== undefined) {
    payload['token'] = token;
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: payload,
    });
    Logger.log(`Web API (files.upload) response: ${response}`);
    return response;
  } else {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      headers: {Authorization: `Bearer ${token}`},
      payload: payload,
    });
    Logger.log(`Web API (files.upload) response: ${response}`);
    return response;
  }
}
