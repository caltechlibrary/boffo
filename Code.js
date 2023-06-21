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

// The order here determines the order of the columns in the results sheet.
// The length of this array also determines the number of columns.
const fields = [
  ['Barcode'                  , (item) => item.barcode                                  ],
  ['Title'                    , (item) => item.title                                    ],
  ['Material type'            , (item) => item.materialType.name                        ],
  ['Status'                   , (item) => item.status.name                              ],
  ['Effective location'       , (item) => item.effectiveLocation.name                   ],
  ['Effective call number'    , (item) => item.effectiveCallNumberComponents.callNumber ],
  ['Enumeration'              , (item) => item.enumeration                              ],
  ['Effective shelving order' , (item) => item.effectiveShelvingOrder                   ],
  ['Item UUID'                , (item) => item.id                                       ],
]

// Regexp for testing that a string looks like a valid Caltech Library barcode.
const barcodePattern = new RegExp('350\\d+|\\d{1,3}|nobarcode\\d+|temp-\\w+|tmp-\\w+|SFL-\\w+', 'i');

// Regexp for testing that a string looks something like a FOLIO tenant id.
const tenantIdPattern = new RegExp('\\d+');


// Menu definition.
// ............................................................................

function onOpen() {
  // Note: the spaces after the icons are actually 2 unbreakable spaces.
  ui.createMenu('Boffo 🐝 ﻿')
    .addItem('🔎 ﻿ ﻿Look up barcodes in FOLIO', 'lookUpBarcodes')
    .addSeparator()
    .addItem('🪪 ﻿ ﻿Set FOLIO credentials', 'getFolioCredentials')
    .addItem('▥ ﻿ ﻿ About Boffo', 'showAbout')
    .addToUi();
}

function onInstall() {
  onOpen();
}



// Functions for getting/setting FOLIO server URL and tenant ID.
// ............................................................................

/**
 * Checks that the stored FOLIO URL and tenant id are valid by contacting the
 * server with a query that should work without needing a token.
 */
function plausibleFolioUrl(url) {
  return url && url.startsWith('https://');
}

/**
 * Does basic sanity-checking on a string to check that it looks like a FOLIO
 * tenant id and not (e.g.) a URL.
 */
function plausibleTenantId(id) {
  return id && !id.startsWith('https') && tenantIdPattern.test(id);
}

/**
 * Returns true if it looks like the necessary data to use the FOLIO API
 * have been stored.
 */
function haveValidCredentials() {
  let url = scriptProps.getProperty('boffo_folio_url');
  let id  = scriptProps.getProperty('boffo_folio_tenant_id');
  if (!plausibleFolioUrl(url) || !plausibleTenantId(id)) {
    return false;
  }

  // The only way to check the token is to try to make an API call.
  let endpoint = url + '/instance-statuses?limit=0';
  let token = userProps.getProperty('boffo_folio_api_token');
  let options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': id,
      'x-okapi-token': token
    },
    'muteHttpExceptions': true
  };
  
  log(`testing if token is valid by doing HTTP get on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let httpCode = response.getResponseCode();
  log(`got response from Folio with HTTP code ${httpCode}`);
  return httpCode < 400;
}

/**
 * Checks that the FOLIO server URL and tenant ID are set to valid-looking
 * values. If they are not, uses a dialog to ask the user for the values and
 * then stores them in the script properties.
 */
function checkFolioCredentials() {
  if (haveValidCredentials()) {
    log('Folio credentials look okay');
  } else {
    log('Folio url, tenant_id, and/or token are not set or are invalid');
    getFolioCredentials();
  }
}

/**
 * Gets FOLIO server URL and tenant ID values from the user and stores them
 * in the script properties. The process is split into two parts: this
 * function creates a dialog using an HTML form, and then JavaScript code
 * embedded in the HTML form calls the separate function saveFolioInfo().
 */
function getFolioCredentials() {
  try {
    let htmlTemplate = HtmlService.createTemplateFromFile('folio-form');
    // Setting the next variable on the template makes it available in the
    // script code embedded in the HTML source of folio-form.html.
    htmlTemplate.waiting = true;
    const htmlContent = htmlTemplate.evaluate().setWidth(475).setHeight(350);
    log('showing dialog to ask user for Folio creds');
    ui.showModalDialog(htmlContent, 'FOLIO information needed');
    // I don't like busy loops but the alternatives look much more complicated.
    let maxWaitSeconds = 60;
    log(`waiting max of ${maxWaitSeconds}s for user to fill in dialog`);
    while (htmlTemplate.waiting && --maxWaitSeconds*2 > 0) {
      Utilities.sleep(500);             // 1/2 sec per loop.
    }
    log(`done; waiting = ${htmlTemplate.waiting}`);
  } catch (err) {
    log('got exception asking for credentials: ' + err);
    quit();
  }
}

/**
 * Gets called by the submit() method in folio-form.html.
 */
function saveFolioInfo(url, tenantId, user, password) {
  // Start with some basic sanity-checking.
  log(`user submitted form with url = ${url}`);
  if (!plausibleFolioUrl(url)) {
    ui.alert("The given URL does not look like a URL and cannot be used.");
    return;
  }
  if (!plausibleTenantId(tenantId)) {
    ui.alert("The given tenant ID does not look valid. It cannot be used.");
    return;
  }

  // Looks good. Save those.
  scriptProps.setProperty('boffo_folio_url', url);
  scriptProps.setProperty('boffo_folio_tenant_id', tenantId);

  // Now try to create the token.
  let endpoint = url + '/authn/login';
  let payload = JSON.stringify({
      'tenant': tenantId,
      'username': user,
      'password': password
  });
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload,
    'headers': {
      'x-okapi-tenant': tenantId
    },
    'muteHttpExceptions': true
  };
  log(`doing HTTP post on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let httpCode = response.getResponseCode();
  log(`got response from Folio with HTTP code ${httpCode}`);
  if (httpCode < 300) {
    let response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      // We have a token!
      let token = response_headers['x-okapi-token'];
      userProps.setProperty('boffo_folio_api_token', token);
      log('got token from Folio and saved it');
    } else {
      ui.alert('Folio did not return a token');
    }
  } else if (httpCode == 422) {
    let results = JSON.parse(response.getContentText());
    let folioMsg = results.errors[0].message;
    let question = `FOLIO rejected the request: ${folioMsg}. Try again?`;
    if (ui.alert(question, ui.ButtonSet.YES_NO) == ui.Button.YES) {
      getFolioCredentials();
    } else {
      ui.alert('Use the menu option "Set FOLIO credentials" to add'
               + ' valid credentials when you are ready. Until then,'
               + ' FOLIO lookup operations will fail.');
    }
  } else {
    ui.alert(`An error occurred communicating with Folio (code ${httpCode}).`);
  }
}


// Functions for looking up info about items.
// ............................................................................

/**
 * Reads the barcodes selected in the current spreadsheet, looks them up in
 * FOLIO, and creates a new sheet with columns containing item field values.
 */
function lookUpBarcodes() {
  // First check that we have valid creds, and ask the user for them if not.
  checkFolioCredentials();
  // If we still don't have valid creds, we bail.
  if (!haveValidCredentials()) {
    ui.alert('Unable to continue due to missing token and/or Folio server info');
    quit();
  }

  // If we get here, we have credentials and we are ready to do the thing.
  try {
    // Get array of values from spreadsheet. This will be a list of strings.
    let selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
    let barcodes = selection.getActiveRange().getDisplayValues();

    // Filter out strings that don't look like barcodes.
    barcodes = barcodes.filter(x => barcodePattern.test(x));
    let numBarcodes = barcodes.length;
    if (numBarcodes < 1) {
      // Either the selection was empty, or filtering removed everything.
      ui.alert('Boffo', 'Please select cells with item barcodes.', ui.ButtonSet.OK);
      return;
    }

    // Create a new sheet where results will be written.
    let resultsSheet = createResultsSheet(fields.map((field) => field[0]));
    let lastLetter = lastColumnLetter();

    // Get item data for each barcode & write to the sheet.
    log(`getting ${numBarcodes} records`);
    for (let i = 0, bc = barcodes[i]; i < numBarcodes; bc = barcodes[++i]) {
      ss.toast(`Looking up ${bc} (item ${i+1} of ${numBarcodes}) …`, 'Boffo', -1);

      let data = itemData(bc);
      let row = i + 2;                  // Offset +1 for header row.
      if (data !== null) {
        let cells = resultsSheet.getRange(`A${row}:${lastLetter}${row}`);
        cells.setValues([fields.map((field) => field[1](data))]);
      } else {
        let cell = resultsSheet.getRange(`A${row}`);
        cell.setValues([bc]);
        cell.setFontColor('red');
      }
    }
    log('done writing data to sheet');
    ss.toast('Done! ✨', 'Boffo', 1);
  } catch (err) {
    log('got exception looking up barcode(s): ' + err);
    quit();
  }  
  return;
}

/**
 * Returns the FOLIO item data for a given barcode.
 */
function itemData(barcode) {
  let url = scriptProps.getProperty('boffo_folio_url');
  let endpoint = url + '/inventory/items?query=barcode=' + barcode;
  let options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': scriptProps.getProperty('boffo_folio_tenant_id'),
      'x-okapi-token': userProps.getProperty('boffo_folio_api_token')
    },
    'muteHttpExceptions': true
  };
  
  log(`doing HTTP post on ${endpoint}`);
  let response = UrlFetchApp.fetch(endpoint, options);
  let httpCode = response.getResponseCode();
  log(`got response from Folio with HTTP code ${httpCode}`);
  // If an error occurred, report it now and stop.
  if (httpCode >= 300) {
    log('alerting user to the error and stopping.');
    switch (httpCode) {
      case 401:
      case 403:
        ui.alert('A FOLIO authentication error occurred. The account'
                 + ' credentials used may be invalid, or the account'
                 + ' may not have the necessary permissions in FOLIO'
                 + ' to perform the action requested. You can try to'
                 + ' reset the FOLIO credentials (use the Boffo menu'
                 + ' option for that). If the error persists, please'
                 + ' contact the FOLIO administrators for assistance.');
        break;
      case 404:
        ui.alert('The API call made by Boffo does not appear to exist at'
                 + ' the address Boffo attempted to use. This may be due'
                 + ' to a temporary network glitch. Please wait a moment'
                 + ' then retry the same operation again. If the problem'
                 + ' persists, please report this to the developers.');
        break;
      case 409:
      case 500:
      case 501:
        ui.alert('FOLIO turned an internal server error. This might be due'
                 + ' to a temporary problem with FOLIO itself. Please wait'
                 + ' a moment, then retry the same operation. If the error'
                 + ' persists, please report it to the developers.'
                 + ` (Error code ${httpCode}.)`);
        break;
      default:
        ui.alert(`An error occurred communicating with FOLIO `
                 + ` (code ${httpCode}). Please report this`
                 + ' to the developers.');
    }
    ss.toast('Stopping due to error.', 'Boffo', 2);
    quit();
  }

  let results = JSON.parse(response.getContentText());
  log(`results for ${barcode}: ` + response.getContentText());
  if (results.totalRecords == 0) {
    log(`Folio did not return data for ${barcode}`);
    return null;
  } else if (results.totalRecords > 1) {
    // FIXME put something in the output
    log(`Folio returned more than one item for ${barcode}`);
    return null;
  } else {
    return results.items[0];
  }
}

/**
 * Creates the results sheet and returns it.
 */
function createResultsSheet(headings) {
  let sheet = ss.insertSheet(uniqueSheetName());
  sheet.setColumnWidths(1, numColumns(), 150);

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


// Functions for showing "about" screen.
// ............................................................................

function showAbout() {
  const htmlTemplate = HtmlService.createTemplateFromFile('about');
  // Setting the next variable on the template makes it available in the
  // script code embedded in the HTML source of about.html.
  htmlTemplate.boffo = getBoffoData();
  const htmlContent = htmlTemplate.evaluate().setWidth(250).setHeight(200);
  log('showing about dialog');
  ui.showModalDialog(htmlContent, 'About Boffo');
}


// Functions used in HTML files.
// ............................................................................
// These are called using the <?!= functioncall(); ?> mechanism in the HTML
// files used in this project.

/**
 * Returns the content of an HTML file in this project. Code originally from
 * https://developers.google.com/apps-script/guides/html/best-practices
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns a scripts property value.
 */
function getProp(prop) {
  if (prop) {
    return scriptProps.getProperty(prop);
  } else {
    log(`called getProp() with an empty string`);
  }
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
 * Returns a JSON object containing fields for the version number and other
 * info about this software. The field names and values on the object returned
 * by this function match exactly the fields in the codemeta.json file.
 */  
function getBoffoData() {
  // Ideally, we would simply read the codemeta.json file. Unfortunately,
  // Google Apps Scripts only provides a way to read HTML files in the local
  // script directory, not JSON files. But that won't stop us! If we add a
  // symlink in the repository named "version.html" pointing to codemeta.json,
  // voilà, we can read it using HtmlService and parse the content as JSON.

  let codemetaFile = {};
  let errorText = 'This installation of Boffo has been damaged somehow:'
      + ' either some files are missing from the installation or one or'
      + ' more files are not in the expected format. Please report this'
      + ' error to the developers.';
  let errorThrown = new Error('Unable to continue.');

  try {  
    codemetaFile = HtmlService.createHtmlOutputFromFile('version.html');
  } catch ({name, message}) {
    log('Unable to read version.html: ' + message);
    ui.alert(errorText);
    throw errorThrown;
  }
  try {
    return JSON.parse(codemetaFile.getContent());
  } catch ({name, message}) {
    log('Unable to parse JSON content of version.html: ' + message);
    ui.alert(errorText);
    throw errorThrown;
  }
}

function log(text) {
  Logger.log(text);
};

function quit() {
  log('quitting');
  throw new Error('Stopping.');
}
