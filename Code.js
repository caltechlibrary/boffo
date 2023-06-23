// ============================================================================
// @file    Code.js (shown as Code.gs in the Google Apps Script editor)
// @brief   Main file for Boffo
// @created 2023-06-08
// @license Please see the file named LICENSE in the project directory
// @website https://github.com/caltechlibrary/boffo
// ============================================================================


// Global constants.
// ............................................................................

// Shortcuts to objects in the Google Apps Script environment.
const ui = SpreadsheetApp.getUi();
const ss = SpreadsheetApp.getActiveSpreadsheet();
const props = PropertiesService.getUserProperties();

// FOLIO item record fields put into the results sheet when the user calls on
// findBarcodes. The order determines the order of the columns in the results
// sheet, and the length of this array determines the number of columns.
const fields = [
  ['Barcode'                  , (item) => item.barcode],
  ['Title'                    , (item) => item.title],
  ['Material type'            , (item) => item.materialType.name],
  ['Status'                   , (item) => item.status.name],
  ['Effective location'       , (item) => item.effectiveLocation.name],
  ['Effective call number'    , (item) => item.effectiveCallNumberComponents.callNumber],
  ['Enumeration'              , (item) => item.enumeration],
  ['Effective shelving order' , (item) => item.effectiveShelvingOrder],
  ['Item UUID'                , (item) => item.id],
]

// Regexp for testing that a string looks like a valid Caltech Library barcode.
const barcodePattern = new RegExp('350\\d+|\\d{1,3}|nobarcode\\d+|temp-\\w+|tmp-\\w+|SFL-\\w+', 'i');


// Google Sheets add-on menu definition.
// ............................................................................
// This creates the "Boffo" menu item in the Extensions menu in Google Sheets.
// Note: onInstall() gets called when the user installs the add-on; onOpen()
// gets called when they open a sheet (or do a page reload in their browser).
// However, onInstall() doesn't seem to get called when you are working on
// Boffo in a script project in Google's online editor, so both functions have
// to be defined, even though they end up doing the same thing.

function onOpen() {
  // BEWARE: in the addItem calls below, the spaces after the icons are
  // actually 2 unbreakable spaces. This detail is invisible in the editor.
  ui.createMenu('Boffo')
    .addItem('🔎 ﻿ ﻿Look up barcodes in FOLIO', 'menuItemLookUpBarcodes')
    .addSeparator()
    .addItem('🪪 ﻿ ﻿Set FOLIO credentials', 'menuItemGetCredentials')
    .addItem(' ✘ ﻿ ﻿ Clear FOLIO token', 'menuItemClearToken')
    .addItem('ⓘ ﻿ ﻿ About Boffo', 'menuItemShowAbout')
    .addToUi();
  log('Boffo menu added');
}

function onInstall() {
  onOpen();

  // We use script properties to pre-populate the FOLIO OKAPI URL and tenant
  // id value fields in the credentials form (folio-form.html). However, the
  // way that the script properties work is that anyone in the org can change
  // the value if they type some other value in our credentials form. So, to
  // avoid that, we copy the values to the user properties and work off that.
  const scriptProps = PropertiesService.getScriptProperties();
  let url = scriptProps.getProperty('boffo_folio_url');
  if (url) {
    props.setProperty('boffo_folio_url', url);
    log(`set user property boffo_folio_url to ${url}`);
  }
  let id  = scriptProps.getProperty('boffo_folio_tenant_id');
  if (id) {
    props.setProperty('boffo_folio_tenant_id', id);
    log(`set user property boffo_folio_tenant_id to ${id}`);
  }
}


// Menu item "Look up barcodes".
// ............................................................................

/**
 * Ensures that the FOLIO credentials are set and then calls lookUpBarcodes().
 */
function menuItemLookUpBarcodes() {
  withCredentials(lookUpBarcodes);
}

/**
 * Reads the barcodes selected in the current spreadsheet, looks them up in
 * FOLIO, and creates a new sheet with columns containing item field values.
 */
function lookUpBarcodes() {
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
    note(`Looking up ${bc} (item ${i+1} of ${numBarcodes}) …`, -1);

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
  note('Done! ✨', 1);
}

/**
 * Returns the FOLIO item data for a given barcode.
 */
function itemData(barcode) {
  let url = props.getProperty('boffo_folio_url');
  let endpoint = url + '/inventory/items?query=barcode=' + barcode;
  let options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': props.getProperty('boffo_folio_tenant_id'),
      'x-okapi-token': props.getProperty('boffo_folio_api_token')
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
    note('Stopping due to error.', 0.1);
    switch (httpCode) {
      case 400:
      case 401:
      case 403:
        quit('Stopped due to an error',
             'A FOLIO authentication error occurred. This can be'
             + ' due to an invalid FOLIO URL, tenant ID, or token,'
             + " or an account that doesn't have the permissions"
             + ' to perform the action requested. You can try to'
             + ' reset the FOLIO credentials (use the Boffo menu'
             + ' option for that). If the error persists, please'
             + ' contact the FOLIO administrator for assistance.');
        break;
      case 404:
        quit('Stopped due to an error',
             'The API call made by Boffo does not appear to exist at'
             + ' the address Boffo attempted to use. This may be due'
             + ' to a temporary network glitch. Please wait a moment'
             + ' then retry the same operation again. If the problem'
             + ' persists, please report this to the developers.');
        break;
      case 409:
      case 500:
      case 501:
        quit('Stopped due to an error',
             'FOLIO turned an internal server error. This might be due'
             + ' to a temporary problem with FOLIO itself. Please wait'
             + ' a moment, then retry the same operation. If the error'
             + ' persists, please report it to the developers.'
             + ` (Error code ${httpCode}.)`);
        break;
      default:
        quit('Stopped due to an error',
             `An error occurred communicating with FOLIO `
             + ` (code ${httpCode}). Please report this`
             + ' to the developers.');
    }
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


// Menu item "Get Credentials" and related functionality.
// ............................................................................

/**
 * Invokes the dialog to set FOLIO credentials and create an API token.
 */
function menuItemGetCredentials() {
  getCredentials();
}

/**
 * Unconditionally shows the dialog to get Folio creds from the user and
 * saves them if a token is successfully obtained.
 */
function getCredentials() {
  let dialog = buildCredentialsDialog();
  log('showing dialog to get credentials');
  ui.showModalDialog(dialog, 'FOLIO credentials');
}

/**
 * Returns an HTML template object for requesting FOLIO creds from the user.
 * If optional argument callAfterSuccess is not empty, it must be the name
 * of a Boffo function to be called after the credentials are successfully
 * used to create a FOLIO token.
 */
function buildCredentialsDialog(callAfterSuccess = '') {
  log(`building form for Folio creds; callAfterSuccess = ${callAfterSuccess}`);
  let htmlTemplate = HtmlService.createTemplateFromFile('folio-form');
  htmlTemplate.callAfterSuccess = callAfterSuccess;
  return htmlTemplate.evaluate().setWidth(475).setHeight(350);
}

/**
 * Returns true if it looks like the necessary data to use the FOLIO API have
 * been stored. This does NOT verify that the credentials actually work;
 * doing so requires an API call, which takes time and slows down user
 * interaction. The approach taken in most functions in Boffo is to only
 * check that the necessary data has been set (by calling this function) and
 * let any exceptions occur at the time commands are executing.
 */
function haveCredentials() {
  return (nonempty(props.getProperty('boffo_folio_url'))
          && nonempty(props.getProperty('boffo_folio_tenant_id'))
          && nonempty(props.getProperty('boffo_folio_api_token')));
}

/**
 * Invokes a function after first making sure FOLIO credentials have been
 * configured. The rationale for doing things this way is the following. When
 * users invoke a command (like lookUpBarcodes) before setting their
 * credentials, it's annoying to them if we stop and tell them to invoke a
 * DIFFERENT menu item first and THEN come back and run the original menu
 * item AGAIN. So for commands that need creds, we wrap them in this
 * function. This gets credentials if needed and then chain-calls the
 * desired function.
 */
function withCredentials(funcToCall) {
  if (haveCredentials()) {
    log(`have credentials; calling ${funcToCall.name}`);
    funcToCall();
  } else {
    log(`need to get credentials before calling ${funcToCall.name}`);
    let dialog = buildCredentialsDialog(funcToCall.name);
    ui.showModalDialog(dialog, 'FOLIO information needed');
  }
}

/**
 * Gets called by the submit() method in folio-form.html. Returns true
 * if successful, false or an exception if not successful.
 *
 * IMPORTANT: this function is called from the context of the client-side
 * Javascript code in folio-form.html. Any exceptions in this function will
 * not be caught by the spreadsheet framework; instead, they'll go to the
 * google.script.run framework's withFailureHandler() handler. This means
 * that calling ui.alert() here works (as long as the code in folio-form.html
 * closes the dialog, which it does in the failure handler) but exceptions
 * will only show up in the logs, without the usual black error banner across
 * the top of the spreadsheet.
 */
function saveFolioInfo(url, tenantId, user, password, callAfterSuccess = '') {
  // Before saving the given info, we try to create a token. If that fails,
  // don't save the input because something is probably wrong with it.
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
  let response = undefined;
  let httpCode = undefined;
  try {
    response = UrlFetchApp.fetch(endpoint, options);
    httpCode = response.getResponseCode();
    log(`got response from Folio with HTTP code ${httpCode}`);
  } catch ({name, message}) {
    log(`error attempting UrlFetchApp on ${endpoint}: ${message}`);
    quit('Failed due to an unrecognized error', 
         'Please carefully check the values entered in the form. For example,'
         + ' is the URL correct and well-formed? Are there any typos'
         + ' anywhere? If you cannot find the problem or it appears to be'
         + ' elsewhere (perhaps a bug in Boffo), please contact the developers'
         + ` and report the following error message:  ${message}`, true);
    return false;
  }

  if (httpCode < 300) {
    let response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      let token = response_headers['x-okapi-token'];
      props.setProperty('boffo_folio_api_token', token);
      log('got token from Folio and saved it');
      // Also save the URL & tenant id now, since we know they work.
      props.setProperty('boffo_folio_url', url);
      props.setProperty('boffo_folio_tenant_id', tenantId);
      return true;
    } else {
      log('no token in the FOLIO response headers');
      quit('Unexpectedly failed to get a token back',
           'The call to FOLIO was successful, but FOLIO did not return'
           + ' a token. This situation should never occur and probably'
           + ' indicates a bug in Boffo. Please report this to the'
           + ' developers and describe what led to it.', true);
      return false;
    }

  } else if (httpCode < 500) {
    let responseContent = response.getContentText();
    let folioMsg = responseContent;
    if (nonempty(responseContent) && responseContent.startsWith('{')) {
      let results = JSON.parse(response.getContentText());
      folioMsg = results.errors[0].message;
    }
    let question = `FOLIO rejected the request: ${folioMsg}. Try again?`;
    if (ui.alert(question, ui.ButtonSet.YES_NO) == ui.Button.YES) {
      // Recursive call.
      log('user chose to try again');
      if (callAfterSuccess) {
        withCredentials(eval(callAfterSuccess));
      } else {
        getCredentials();
      }
      return haveCredentials();
    } else {
      quit("Stopped at the user's request",
           'You can use the menu option "Set FOLIO credentials" to'
           + ' add valid credentials when you are ready. Until then,'
           + ' FOLIO lookup operations will fail.', true);
      return false;
    }

  } else {
    quit('Failed due to a FOLIO server or network problem',
         'This may be temporary. Try again after waiting a short time. If'
         + ' the error persists, please contact the FOLIO administrators'
         + ' and/or the developers of Boffo. (When reporting  the error,'
         + ` please mention this was an HTTP code ${httpCode} error.)`, true);
    return false;
  }

  return haveCredentials();
}

/**
 * Runs a function given its name as a string. This is invoked by the code
 * in folio-form.html and is needed because we can't pass a function object
 * to the form. If we could, the form could invoke that function directly;
 * instead, the form runs google.script.run.callBoffoFunction(name). Note
 * that this function calls JavaScript eval, which is considered bad, but
 * in this situation, we control what is put into the form, and thus it is
 * safe to do it this way.
 */
function callBoffoFunction(name) {
  log(`running ${name}()`);
  eval(name)();
}

/**
 * Returns true if the user's stored token is valid, based on contacting FOLIO.
 * Note: not currently used, but kept in case it's needed in the future.
 */
function haveValidToken() {
  let url   = props.getProperty('boffo_folio_url');
  let id    = props.getProperty('boffo_folio_tenant_id');
  let token = props.getProperty('boffo_folio_api_token');

  // The only way to check the token is to try to make an API call.
  let endpoint = url + '/instance-statuses?limit=0';
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
  log(`got code ${httpCode}; token ${httpCode < 400 ? "is" : "is not"} valid`);
  return httpCode < 400;
}


// Menu item "Clear token".
// ............................................................................

/**
 * Unsets the user's stored token. This should rarely be needed but might be
 * useful for debugging.
 */
function menuItemClearToken() {
  log('deleting stored token');
  props.deleteProperty('boffo_folio_api_token');
  ui.alert('Your stored FOLIO token has been deleted. You can use the'
           + ' menu option "Set FOLIO credentials" to generate a new one.');
}


// Menu item "About Boffo".
// ............................................................................

function menuItemShowAbout() {
  const htmlTemplate = HtmlService.createTemplateFromFile('about');
  // Setting the next variable on the template makes it available in the
  // script code embedded in the HTML source of about.html.
  htmlTemplate.boffo = getBoffoMetadata();
  const htmlContent = htmlTemplate.evaluate().setWidth(250).setHeight(200);
  log('showing about dialog');
  ui.showModalDialog(htmlContent, 'About Boffo');
}


// Helper functions used in HTML scriptlets.
// ............................................................................
// These are invoked within our HTML templates using Google's "scriptlet"
// feature, which are constructs of the form <?= functioncall(); ?> and
// <?!= functioncall(); ?>.  The Google docs for this can be found at
// https://developers.google.com/apps-script/guides/html/templates

/**
 * Returns the content of an HTML file in this project. Code originally from
 * https://developers.google.com/apps-script/guides/html/best-practices
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns a user property value.
 */
function getProp(prop) {
  if (prop) {
    return props.getProperty(prop);
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
function getBoffoMetadata() {
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

/**
 * Shows a brief note to the user. This is a separate function instead
 * of simply calling SpreadsheetApp.getActiveSpreadsheet().toast(...)
 * so that we can invoke it from folio-form.html. The default duration
 * is 2 seconds.
 */
function note(message, duration = 2) {
  ss.toast(message, 'Boffo', duration);
}

/**
 * Quits execution by throwing an Error, which causes Google to display
 * a black error banner across the top of the page.
 */
function quit(why, details = '', showAlert = false) {
  // Custom exception object. This is used so that the Google error banner
  // says "Why: details" instead of the default ("Error: ...message...").
  function Stopped() {
    this.name = why;
    this.message = details;
  };
  Stopped.prototype = Error.prototype;
  if (showAlert) {
    ui.alert(why + '. ' + details);
  }
  throw new Stopped();
}

/**
 * Returns true if the given value is not empty, null, or undefined.
 */
function nonempty(value) {
  return value ? true : false;
}

function log(text) {
  Logger.log(text);
}
