var authToken = null;

function makeApiPostCall_(methodUrl, json_payload)
{
  if (authToken == null) authToken = props.getProperty("token");
  var url = 'https://api-invest.tinkoff.ru/openapi/' + methodUrl;  
  var response = UrlFetchApp.fetch(url, {
    'headers': {'accept': 'application/json', "Authorization": `Bearer ${authToken}`},
    'method' : 'post',
    'escaping': false,
    'muteHttpExceptions':true,
    "payload" : json_payload});
  
  if (response.getResponseCode() != 200) return null;
  //Logger.log(response.getContentText()); 
  return JSON.parse(response.getContentText()).payload;
}

function makeApiGetCall_(methodUrl)
{
  if (authToken == null) authToken = props.getProperty("token");
  var url = 'https://api-invest.tinkoff.ru/openapi/' + methodUrl;
  var response = UrlFetchApp.fetch(url, {
    'headers': {'accept': 'application/json', "Authorization": `Bearer ${authToken}`},
    'escaping': false,
    'muteHttpExceptions':true});
  
  if (response.getResponseCode() != 200) return null;
  //Logger.log(response.getContentText());
  return JSON.parse(response.getContentText()).payload;
}
