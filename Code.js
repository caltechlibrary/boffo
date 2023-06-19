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

// Regexp for testing that a string looks something like a FOLIO tenant id.
const tenantIdPattern = new RegExp('\\d+');

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

const helpURL = 'http://caltechlibrary.github.io/boffo';
const cancel = new Error('cancelled');


// Menu definition.
// ............................................................................

function onOpen() {
  // Note: the spaces after the icons are actually 2 unbreakable spaces.
  ui.createMenu('Boffo 🐝 ﻿')
    .addItem('🔎 ﻿ ﻿Look up barcodes in FOLIO', 'lookUpBarcodes')
    .addSeparator()
    .addItem('🪪 ﻿ ﻿Set FOLIO credentials', 'getFolioCredentials')
    .addItem('⁇ ﻿ ﻿Get help', 'getHelp')    
    .addToUi();
}

function onInstall() {
  onOpen();
}


// Functions for getting/setting FOLIO server URL and tenant ID.
// ............................................................................

/**
 * Does basic sanity-checking on a string to check that it looks like a URL.
 */
function validFolioUrl(url) {
  return url && url.startsWith('https://');
}

/**
 * Does basic sanity-checking on a string to check that it looks like a FOLIO
 * tenant id and not (e.g.) a URL.
 */
function validTenantId(id) {
  return id && tenantIdPattern.test(id) && !id.startsWith('https');
}

/**
 * Does basic sanity-checking on a string to check that it looks like it could
 * be a FOLIO API token. 
 */
function validAPIToken(token) {
  return token && token.length > 150;
}

/**
 * Returns true if it looks like the necessary data to use the FOLIO API
 * have been stored.
 */
function haveFolioCredentials() {
  let url   = scriptProps.getProperty('boffo_folio_url');
  let id    = scriptProps.getProperty('boffo_folio_tenant_id');
  let token = userProps.getProperty('boffo_folio_api_token');
  return validFolioUrl(url) && validTenantId(id) && validAPIToken(token);
}

/**
 * Checks that the FOLIO server URL and tenant ID are set to valid-looking
 * values. If they are not, uses a dialog to ask the user for the values and
 * then stores them in the script properties.
 */
function checkFolioCredentials() {
  if (haveFolioCredentials()) {
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
    let htmlContent = HtmlService
        .createTemplateFromFile('folio-form')
        .evaluate()
        .setWidth(475)
        .setHeight(330);
    log('showing dialog to ask user for Folio URL & tenant id');
    ui.showModalDialog(htmlContent, 'FOLIO information needed');
  } catch (err) {
    quit(err);
  }
}

/**
 * Gets called by the submit() method in folio-form.html.
 */
function saveFolioInfo(url, tenant_id, user, password) {
  // Start with some basic sanity-checking.
  log(`user submitted form with url = ${url}`);
  if (!validFolioUrl(url)) {
    ui.alert("The given URL does not look like a URL and cannot be used.");
    return;
  }
  if (!validTenantId(tenant_id)) {
    ui.alert("The given tenant ID looks like a URL instead. It cannot be used.");
    return;
  }

  // Looks good. Save those.
  scriptProps.setProperty('boffo_folio_url', url);
  scriptProps.setProperty('boffo_folio_tenant_id', tenant_id);

  // Now try to create the token.
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
      // We have a token!
      let token = response_headers['x-okapi-token'];
      userProps.setProperty('boffo_folio_api_token', token);
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

/**
 * Reads the barcodes selected in the current spreadsheet, looks them up in
 * FOLIO, and creates a new sheet with columns containing item field values.
 */
function lookUpBarcodes() {
  // Check if we have creds, ask user for them if we not, and if we don't
  // end up getting the values, bail.
  checkFolioCredentials();
  if (!haveFolioCredentials()) {
    ui.alert('Unable to continue due to missing token and/or Folio server info');
    return;
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
    quit(err);
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


// Functions for showing help.
// ............................................................................
// The approach used here is based on a blog posting by Amit Agarwal from
// 2022-05-07 at https://www.labnol.org/open-webpage-google-sheets-220507

/**
 * Opens a window on the Boffo help pages if possible, or opens a dialog
 * that asks the user to click a link to get to the help pages.
 */
function getHelp() {
  const htmlTemplate = HtmlService.createTemplateFromFile('help');
  // Setting the next variable on the template makes it available in the
  // script code embedded in the HTML source of help.html.
  htmlTemplate.boffoHelpURL = helpURL;
  const htmlContent = htmlTemplate.evaluate().setWidth(260).setHeight(180);
  log('showing help dialog');
  ui.showModalDialog(htmlContent, 'Help for Boffo');
  Utilities.sleep(1500);
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

function log(text) {
  Logger.log(text);
};

function quit(msg) {
  log(`stopped execution: ${msg}`);
}
