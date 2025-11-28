const MODEL_ID = 'gemini-1.0-pro';
const PROJECT_ID = 'agentbuilder-429502';
const REGION = 'us-central1';
const API_ENDPOINT = `https://${REGION}-aiplatform.googleapis.com/v1/projects/${PROJECT_ID}/locations/${REGION}/publishers/google/models/${MODEL_ID}:streamGenerateContent`;

function onOpen() {
  let ui = DocumentApp.getUi();

  ui.createMenu('Gemini')
      .addItem('Authorise', 'auth')
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  let ui = DocumentApp.getUi();
  let promptFrom = HtmlService.createHtmlOutputFromFile('index').setTitle('Help Me Write');

  ui.showSidebar(promptFrom);
}

function auth() {
  cache = CacheService.getUserCache();
  token = ScriptApp.getOAuthToken();
  console.log(token);
  cache.put('token', token);
}

function askGemini(prompt) {
  cache = CacheService.getUserCache();
  token = cache.get('token');

  let data = {
    contents: {
      role: 'USER',
      parts: { 'text': prompt }
    }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',   
    headers: {
      Authorization: `Bearer ${token}`,
    },
    payload: JSON.stringify(data)
  };

  console.log(`Token = ${token}`);
  if (!token) {
    insertText('Please generate a new authorisation token: Gemini > Authorise');
    return;
  }
  
  const response = UrlFetchApp.fetch(API_ENDPOINT, options);

  if (response.getResponseCode() == 200) {
    let json = JSON.parse(response.getContentText());
    let answer = '';

    console.log(response.getContentText());
    // TODO: loop through all parts of response for longform answers

    for (let i in json){
      answer += json[i].candidates[0].content.parts[0].text;
    }
    insertText(answer);
    return;
  }
  
  insertText('There was an error. Try generating a new authorisation token: Gemini > Authorise');
  return;
}

function insertText(text) {
  let doc = DocumentApp.getActiveDocument();
  let body = doc.getBody();

  body.appendParagraph(text);
}