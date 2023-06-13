// ============================================================================
// @file    Code.js
// @brief   Main file for Boffo
// @created 2023-06-08
// @license Please see the file named LICENSE in the project directory
// @website https://github.com/caltechlibrary/boffo
// ============================================================================


// Shortcuts to objects in the Google Apps Script environment.
// ............................................................................

const ui = SpreadsheetApp.getUi();
const userProps = PropertiesService.getUserProperties();
const scriptProps = PropertiesService.getScriptProperties();


// Global constants.
// ............................................................................

const barcodePattern = new RegExp('350\\d+|\\d{1,3}|nobarcode\\d+|temp-\\w+|tmp-\\w+|SFL-\\w+', 'i');
const cancel = new Error('cancelled');


// Menu definition.
// ............................................................................

function onOpen() {
  ui.createMenu('FOLIO')
      .addItem('Look up barcodes', 'lookUpBarcode')
      .addSeparator()
      .addItem('Set FOLIO credentials', 'setFolioToken')
      .addItem('Look up item barcodes', 'lookUpBarcode')    
      .addItem('test form', 'showCredentialsForm')
      .addItem('test folio', 'checkFolioServerInfo')    
      .addToUi();
}


// Functions for getting/setting FOLIO server URL and tenant ID.
// ............................................................................

function validFolioUrl(url) {
  return url.startsWith('https://');
}

function validTenantId(tenant_id) {
  return !tenant_id.startsWith('https');
}

function haveFolioServerInfo() {
  let url = scriptProps.getProperty('folio_url');
  let tenant_id = scriptProps.getProperty('tenant_id');
  return validFolioUrl(url) && validTenantId(tenant_id);
}

/**
 * Checks that the FOLIO server URL and tenant ID are set to valid-looking
 * values. If they are not, uses a dialog to ask the user for the values and
 * then stores them in the script properties.
 */
function checkFolioServerInfo() {
  if (haveFolioServerInfo()) {
    log('Folio url and tenant_id look okay');
  } else {
    log('Folio url and/or tenant_id not set or are invalid');
    setFolioServerInfo();
  }
}

/**
 * Gets FOLIO server URL and tenant ID values from the user and stores them
 * in the script properties. The process is split into two parts: this
   function creates a dialog using an HTML form, and then JavaScript code
   embedded in the HTML form calls the separate function saveFolioInfo().
 */
function setFolioServerInfo() {
  try {
    let htmlContent = HtmlService
        .createTemplateFromFile('folio-form')
        .evaluate()
        .setWidth(450)
        .setHeight(240);
    log('asking user for Folio URL & tenant id');
    ui.showModalDialog(htmlContent, 'FOLIO server information');
  } catch (err) {
    quit(err);
  }
}

function saveFolioServerInfo(url, tenant_id) {
  if (!validFolioUrl(url)) {
    ui.alert("The given URL does not look like a URL and cannot be used.");
    return;
  }
  if (!validTenantId(tenant_id)) {
    ui.alert("The given tenant ID looks like a URL instead. It cannot be used.");
    return;
  }
  scriptProps.setProperty('folio_url', url);
  scriptProps.setProperty('tenant_id', tenant_id);
}


// Functions for getting/setting FOLIO API token.
// ............................................................................

function haveToken() {
  return userProps.getProperty('token') != null;
}

function setToken() {
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
    checkFolioServerInfo();
    createNewToken();
  } catch (err) {
    quit(err);
  }
}

/**
 * Asks FOLIO for a new API token. The process is split into two parts: this
 * function creates a dialog using an HTML form, and then JavaScript code
 * embedded in the HTML form calls the separate function getNewToken().
 */
function createNewToken() {
  // The process is split into two parts: this function creates a dialog using
  // an HTML form and JavaScript code embedded in the form calls getNewToken().
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
  
  log(`doing HTTP post on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let http_code = response.getResponseCode();
  log(`got response from Folio with HTTP code ${http_code}`);
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
    ui.alert(`An error occurred communicating with Folio (code ${http_code}).`);
  }
}

function saveToken(token) {
  userProps.setProperty('token', token);
  log('saved token');
}


// Functions for looking up info about items.
// ............................................................................

function lookUpBarcode() {
  // Check if we have creds & ask user for them if we don't.
  if (!haveToken()) {
    setToken();
  }
  if (!haveToken() || !haveFolioServerInfo()) {
    ui.alert('Unable to continue due to missing token and/or Folio server info');
    return;
  }

  try {
    // Get array of values from spreadsheet. This will be a list of strings.
    let selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
    let values = selection.getActiveRange().getDisplayValues();

    // Filter out things that don't look like barcodes.
    values = values.filter(x => barcodePattern.test(x));

    // Ask Folio about each barcode.
    itemData(values[0]);

    // Add column headers to spreadsheet if necessary.
    // Fill out each row with data in the appropriate columns.

  } catch (err) {
    quit(err);
  }  
  // Get the selected cells in the current sheet.
  // Check the values look like barcodes. Ignore ones that don't,
  // and if none are barcodes, raise an alert.
}

function itemData(barcode) {
  let url = scriptProps.getProperty('folio_url');
  let endpoint = url + '/inventory/items?query=barcode=' + barcode;
  let options = {
    'url': endpoint,
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': scriptProps.getProperty('tenant_id'),
      'x-okapi-token': userProps.getProperty('token')
    }
  }
  
  log(`doing HTTP post on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let http_code = response.getResponseCode();
  log(`got response from Folio with HTTP code ${http_code}`);
  if (http_code >= 300) {
    ui.alert(`An error occurred communicating with Folio (code ${http_code}).`);
    return;
  }

  let results = JSON.parse(response.getContentText());
  if (results.totalRecords < 0) {
    log(`Folio did not return data for ${barcode}`);
    return;
  } else if (results.totalRecords > 1) {
    // FIXME put something in the output
    log(`Folio returned more than one item for ${barcode}`);
    return;
  } else {
    let item = results.items[0];
    let id = item.id;
    let title = item.title;
    let call_num = item.callNumber;
    let effective_loc = item.effectiveLocation.name;
    let status = item.status.name;

    log(id);
    log(title);
    log(call_num);
    log(effective_loc);
    log(status);
  }
}


// Miscellaneous helper functions.
// ............................................................................

// Used in the forms html files.
// Originally from https://developers.google.com/apps-script/guides/html/best-practices
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function log(text) {
  Logger.log(text);
};

function quit(msg) {
  log('stopped execution: ' + msg);
}
