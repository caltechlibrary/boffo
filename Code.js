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
const ss = SpreadsheetApp.getActiveSpreadsheet();
const userProps = PropertiesService.getUserProperties();
const scriptProps = PropertiesService.getScriptProperties();


// Global constants.
// ............................................................................

// Regexp for testing that a string looks like a valid Caltech Library barcode.
const barcodePattern = new RegExp('350\\d+|\\d{1,3}|nobarcode\\d+|temp-\\w+|tmp-\\w+|SFL-\\w+', 'i');

// The order here determines the order of the columns in the results sheet.
// The length of this array also determines the number of columns.
const fields = [
  ['Barcode'            , (item) => item.barcode                ],
  ['Title'              , (item) => item.title                  ],
  ['Call number'        , (item) => item.callNumber             ],
  ['Effective location' , (item) => item.effectiveLocation.name ],
  ['Status'             , (item) => item.status.name            ],
  ['UUID'               , (item) => item.id                     ],
]

const helpURL = 'http://caltechlibrary.github.io/boffo';
const cancel = new Error('cancelled');


// Menu definition.
// ............................................................................

function onOpen() {
  // Note: the spaces after the icons are actually 2 unbreakable spaces.
  ui.createMenu('Boffo 🐝 ﻿')
    .addItem('🔎 ﻿ ﻿Look up barcodes in FOLIO', 'lookUpBarcode')
    .addSeparator()
    .addItem('🪪 ﻿ ﻿Set FOLIO credentials', 'createNewToken')
    .addItem('⁇ ﻿ ﻿Get help', 'getHelp')    
    .addToUi();
}

function onInstall() {
  onOpen();
}


// Functions for getting/setting FOLIO server URL and tenant ID.
// ............................................................................

function validFolioUrl(url) {
  return url && url.startsWith('https://');
}

function validTenantId(tenant_id) {
  return tenant_id && !tenant_id.startsWith('https');
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
    log('showing dialog to ask user for Folio URL & tenant id');
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
  let endpoint = url + '/authn/login';
  let payload = JSON.stringify({
      'tenant': tenant_id,
      'username': user,
      'password': password
  });
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload,
    'headers': {
      'x-okapi-tenant': tenant_id
    }
  };
  
  log(`doing HTTP post on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let http_code = response.getResponseCode();
  log(`got response from Folio with HTTP code ${http_code}`);
  if (http_code < 300) {
    let response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      let token = response_headers['x-okapi-token'];
      userProps.setProperty('token', token);
      log('got token from Folio and saved it');
    } else {
      ui.alert('Folio did not return a token');
    }
  } else {
    ui.alert(`An error occurred communicating with Folio (code ${http_code}).`);
  }
}


// Functions for looking up info about items.
// ............................................................................

function lookUpBarcode() {
  // Check if we have creds & ask user for them if we don't.
  checkFolioServerInfo();
  if (!haveToken()) {
    setToken();
  }

  // If we still don't have valid creds, something is wrong & we bail.
  if (!haveToken() || !haveFolioServerInfo()) {
    ui.alert('Unable to continue due to missing token and/or Folio server info');
    return;
  }

  // We've got creds & we're ready to rock.
  try {
    // Get array of values from spreadsheet. This will be a list of strings.
    let selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
    let values = selection.getActiveRange().getDisplayValues();

    // Filter out strings that don't look like barcodes.
    values = values.filter(x => barcodePattern.test(x));
    if (values.length < 1) {
      // Either the selection was empty, or filtering removed everything.
      ui.alert('Boffo', 'Please select some cells containing item barcodes.',
               ui.ButtonSet.OK);
      return;
    }

    // Get item data for each barcode.
    let items = values.map((value) => itemData(value));

    // Create a new sheet and write the data into it.
    let resultsSheet = createResultsSheet(fields.map((field) => field[0]));
    let lastLetter = lastColumnLetter();
    for (let i = 0; i < items.length; i++) {
      let row = i + 2;                  // Offset +1 for header row.
      log(`row = ${row}`);
      let cells = resultsSheet.getRange(`A${row}:${lastLetter}${row}`);
      let data = fields.map((field) => field[1](items[i]));
      log(`data = ${data}`);
      cells.setValues([data]);
    };
    log('done writing data to sheet');
    ss.toast('Done! ✨', 'Boffo', 1);
  } catch (err) {
    quit(err);
  }  
  return;
}

function itemData(barcode) {
  ss.toast(`Getting data for ${barcode} …`, 'Boffo', -1);

  let url = scriptProps.getProperty('folio_url');
  let endpoint = url + '/inventory/items?query=barcode=' + barcode;
  let options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': scriptProps.getProperty('tenant_id'),
      'x-okapi-token': userProps.getProperty('token')
    }
  };
  
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
    log(`returning result for ${barcode}`);
    return results.items[0];
  }
}

function createResultsSheet(headings) {
  let sheet = ss.insertSheet(uniqueSheetName());
  sheet.setColumnWidths(1, numColumns(), 130);

  // FIXME 1000 is arbitrary, picked because new Google sheets have 1000 rows,
  // but it's conceivable someone will create a sheet with more.
  let cells = sheet.getRange('A1:A1000');
  cells.setHorizontalAlignment('left');

  let lastLetter = lastColumnLetter();
  let headerRow = sheet.getRange(`A1:${lastLetter}1`);
  headerRow.setValues([headings]);
  headerRow.setFontSize(10);
  headerRow.setFontColor('white');
  headerRow.setFontWeight('bold');
  headerRow.setBackground('#999999');

  return sheet;
}


// Functions for showing help.
// ............................................................................

function getHelp() {
  let htmlContent = HtmlService
    .createTemplateFromFile('help')
    .evaluate()
    .setWidth(300)
    .setHeight(200);
  log('showing help dialog');
  ui.showModalDialog(htmlContent, 'Help for Boffo');
}

function getHelpURL() {
  return helpURL;
}


// Miscellaneous helper functions.
// ............................................................................

/**
 * Returns the number of columns needed to hold the fields fetched from FOLIO.
 */
function numColumns() {
  return fields.length;
}

/**
 * Returns the spreadsheet column letter corresponding to the last column.
 */
function lastColumnLetter() {
  return 'ABCDEFGHIJKLMNOPQRSTUVWXY'.charAt(fields.length - 1);
}

/**
 * Returns a unique sheet name. The name is generated by taking a base name
 * and, if necessary, appending an integer that is incremented until the name
 * is unique.
 */
function uniqueSheetName(baseName = 'Item Data') {
  let names = ss.getSheets().map((sheet) => sheet.getName());

  // Compare candidate name against existing sheet names & increment counter
  // until we no longer get a match against any existing name.
  let newName = baseName;
  for (let i = 2; names.indexOf(newName) > -1; i++) {
    newName = `${baseName} ${i}`;
  }
  return newName;
}

/**
 * Returns the content of an HTML file in this project. This is used in
 * the HTML files themselves. Code originally based on
 * https://developers.google.com/apps-script/guides/html/best-practices
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function log(text) {
  Logger.log(text);
};

function quit(msg) {
  log(`stopped execution: ${msg}`);
}
