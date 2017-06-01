var userProperties = PropertiesService.getUserProperties();
var documentProperties = PropertiesService.getDocumentProperties();

function onInstall(e) {
  onOpen(e);
  apiPrompt();
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createAddonMenu();
  menu.addItem('Put URLs', 'putUrls');
  menu.addItem('Configuration', 'showSidebar');
  menu.addToUi();
}

function getExtractorId() {
  return documentProperties.getProperty('EXTRACTOR_ID');
}

function getUrlRange() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.setActiveSelection(documentProperties.getProperty('URL_RANGE'));
  range = documentProperties.getProperty('URL_RANGE');
}

function updateUrlRange(inputRange) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var urlRange = sheet.getRange(inputRange);
  sheet.setActiveSelection(urlRange);
  documentProperties.setProperty('URL_RANGE', urlRange);
}

function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('sidebar').evaluate()
  .setTitle('Import.io Configuration');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function apiPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Import.io API Key',
      'Please enter your Import.io API Key:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var api_text = result.getResponseText();
  if (button == ui.Button.OK) {
    userProperties.setProperty('IMPORT_IO_API_KEY', api_text);
    ui.alert('Your API Key has been updated.');
  } else if (button == ui.Button.CANCEL) {
    ui.alert('You will be unable to link to Import.io without your API key.');
  } else if (button == ui.Button.CLOSE) {
    ui.alert('You will be unable to link to Import.io without your API key.');
  }
}

function extractorPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Import.io Extractor ID',
      'Please enter your Import.io Extractor ID:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var extractor_text = result.getResponseText();
  if (button == ui.Button.OK) {
    documentProperties.setProperty('EXTRACTOR_ID', extractor_text);
    ui.alert('The Extractor ID has been updated.');
  } else if (button == ui.Button.CANCEL) {
    ui.alert('You will be unable to link to Import.io without your Extractor ID.');
  } else if (button == ui.Button.CLOSE) {
    ui.alert('You will be unable to link to Import.io without your Extractor ID.');
  }
}

function putUrls() {
  apiKey = userProperties.getProperty('IMPORT_IO_API_KEY');
  extractorId = documentProperties.getProperty('EXTRACTOR_ID');
  Logger.log('API Key: ' + apiKey);
  Logger.log('Extractor ID: ' + extractorId);
  if (apiKey == '') {
    apiPrompt();
  } else if (extractorId == '') {
    extractorPrompt();
  } else {
    urls = getUrls();
    resp = putResp(urls);
    Logger.log('Response: ' + resp)
    finishAlert();
  }
}

function finishAlert() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
     'Your URLs have been placed in the extractor.',
      ui.ButtonSet.OK);
}

function putResp(data) {
  var put_url = 'https://store.import.io/store/extractor/' + documentProperties.getProperty('EXTRACTOR_ID') + '/_attachment/urlList?_apikey=' + userProperties.getProperty('IMPORT_IO_API_KEY');
  var options = {
    'method' : 'put',
    'contentType': 'text/plain',
    'payload' : data
  };
  var resp = UrlFetchApp.fetch(put_url, options);
  return resp
}

function getUrls() {
  var urls = ""
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    urls += data[i][0].replace(/ /g,"%20") + '\n';
  }
  Logger.log('URL List: '+ urls)
  return urls
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}