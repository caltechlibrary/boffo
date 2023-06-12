const ui = SpreadsheetApp.getUi();
const userProps = PropertiesService.getUserProperties();
const scriptProps = PropertiesService.getScriptProperties();

const cancel = new Error('cancelled');
const log = (text) => { Logger.log(text); };

function onOpen() {
  ui.createMenu('FOLIO')
      .addItem('Look up barcodes', 'lookUpBarcode')
      .addSeparator()
      .addItem('Reset FOLIO token', 'resetToken')
      .addItem('form', 'showCredentialsForm')
      .addItem('folio', 'checkFolioInfo')    
      .addToUi();
}

function checkFolioInfo() {
  try {
    if (!scriptProps.getProperty('folio_url') || !scriptProps.getProperty('tenant_id')) {
      // The process is split into two parts: this function creates a dialog using
      // an HTML form, and JavaScript code embedded in the form calls saveFolioInfo().
      let htmlContent = HtmlService
        .createTemplateFromFile('folio-form')
        .evaluate()
        .setWidth(450)
        .setHeight(240);
      log('asking user for Folio URL & tenant id');
      ui.showModalDialog(htmlContent, 'FOLIO server information');
    }
  } catch (err) {
    quit(err);
  }
}

function saveFolioInfo(url, tenant_id) {
  if (!validFolioUrl(url)) {
    ui.alert("The URL value you provided does not look like a URL. It cannot be used.");
    return;
  }
  if (!validTenantId(tenant_id)) {
    ui.alert("The tenant ID should not be a URL.");
    return;
  }
  scriptProps.setProperty('folio_url', url);
  scriptProps.setProperty('tenant_id', tenant_id);
}

function resetToken() {
  try {
    if (haveToken()) {
      log('found an existing token -- asking user if should use it');
      let question = 'A token has already been stored. Generate a new one?';
      if (ui.alert(question, ui.ButtonSet.YES_NO) == ui.Button.NO) {
        log('user said to use existing token');
        return;
      } else {
        log('user said to create new token');
      }
    }
    // Didn't find an existing token, or user said to regenerate it.
    checkFolioInfo();
    getTokenFromFolio();
  } catch (err) {
    quit(err);
  }
}

function haveToken() {
  return userProps.getProperty('token') != null;
}

function saveToken(token) {
  userProps.setProperty('token', token);
  log('saved token');
}

function getTokenFromFolio() {
  // The process is split into two parts: this function creates a dialog using
  // an HTML form, and JavaScript code embedded in the form calls getNewToken().
  let htmlContent = HtmlService
    .createTemplateFromFile('user-form')
    .evaluate()
    .setWidth(450)
    .setHeight(240);
  log('showing credentials form & handing off control to getNewToken()');
  ui.showModalDialog(htmlContent, 'FOLIO user name and password');
}

function getNewToken(user, password) {
  let url = scriptProps.getProperty('folio_url');
  let tenant_id = scriptProps.getProperty('tenant_id');
  if (validFolioUrl(url) && validTenantId(tenant_id)) {
    log('Folio url and tenant_id look okay');
  } else {
    log('stored Folio url and/or tenant_id are invalid; aborting');
    return;
  }
  let endpoint = url + '/authn/login'
  let options = {
    'url': endpoint,
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({
      'tenant': tenant_id,
      'username': user,
      'password': password
    }),
    'headers': {
      'x-okapi-tenant': tenant_id
    }
  }
  
  log('doing HTTP post on ' + endpoint);
  let response = UrlFetchApp.fetch(endpoint, options);
  let http_code = response.getResponseCode();
  log('got response from Folio with HTTP code ' + http_code);
  if (http_code < 300) {
    let response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      let token = response_headers['x-okapi-token'];
      log('got token from Folio');
      saveToken(token);
    } else {
      ui.alert('Folio did not return a token');
    }
  } else {
    ui.alert('An error occurred communicating with Folio (code ' + http_code + ').')
  }
}

function lookUpBarcode() {
  try {
    // Check if we have creds & ask user for them if we don't.
    if (! haveToken()) {
      resetToken();
    }
    // Get array of values from spreadsheet.
    // Filter out things that don't look like barcodes.
    // Ask Folio about each barcode.
    // Add column headers to spreadsheet if necessary.
    // Fill out each row with data in the appropriate columns.

  } catch (err) {
    quit(err);
  }  
  // Get the selected cells in the current sheet.
  // Check the values look like barcodes. Ignore ones that don't,
  // and if none are barcodes, raise an alert.
}

function quit(msg) {
  log('stopped execution: ' + msg);
}

function validFolioUrl(url) {
  return url.startsWith('https://');
}

function validTenantId(tenant_id) {
  return !tenant_id.startsWith('https');
}

// Used in the forms html files.
// Originally from https://developers.google.com/apps-script/guides/html/best-practices
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
