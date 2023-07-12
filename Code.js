// @file    Code.js (shown as Code.gs in the Google Apps Script editor)
// @brief   Main file for Boffo
// @created 2023-06-08
// @license Please see the file named LICENSE in the project directory
// @website https://github.com/caltechlibrary/boffo


// Global constants.
// ............................................................................

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
  ['Item UUID'                , (item) => item.id],
];

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
  SpreadsheetApp.getUi().createMenu('Boffo')
    .addItem('🔎 ﻿ ﻿Look up barcodes in FOLIO', 'menuItemLookUpBarcodes')
    .addSeparator()
    .addItem('🪪 ﻿ ﻿Set FOLIO credentials', 'menuItemGetCredentials')
    .addItem(' ✘ ﻿ ﻿ Clear FOLIO token', 'menuItemClearToken')
    .addItem('ⓘ ﻿ ﻿ About Boffo', 'menuItemShowAbout')
    .addToUi();
}

function onInstall() {
  onOpen();

  // Script properties are used to pre-populate the FOLIO OKAPI URL and
  // tenant id field values in the credentials form (credentials-form.html).
  // This makes it possible for all users in our organization to get pre-filled
  // values for the URL and tenant id without hardwiring values into this
  // code. However, the way that the script properties in a Google Apps
  // Script project work is that anyone in the org can set the script
  // property values. If we were to *only* read/write the properties from/to
  // the script properties, then if any user typed some other values in the
  // credentials form at any point and this program saved the values back to
  // the script properties, it would change the pre-filled values that Boffo
  // users would get after that point. To avoid that, the following code
  // copies the values from the script properties (if they're set) to the
  // user properties (if they don't already exist there), and the rest of the
  // program always works off the user props.
  const scriptProps = PropertiesService.getScriptProperties();
  const userProps = PropertiesService.getUserProperties();
  let url = scriptProps.getProperty('boffo_folio_url');
  if (url && !userProps.getProperty('boffo_folio_url')) {
    url = stripTrailingSlash(url);
    userProps.setProperty('boffo_folio_url', url);
    log(`set user property boffo_folio_url to ${url}`);
  }
  let id = scriptProps.getProperty('boffo_folio_tenant_id');
  if (id && !userProps.getProperty('boffo_folio_tenant_id')) {
    userProps.setProperty('boffo_folio_tenant_id', id);
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
  // Get array of values from spreadsheet. The list we get back is a list of
  // single-item lists, so we also flatten it.
  let selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  let barcodes = selection.getActiveRange().getDisplayValues().flat();

  // Filter out strings that don't look like barcodes.
  barcodes = barcodes.filter(x => barcodePattern.test(x));
  let numBarcodes = barcodes.length;
  if (numBarcodes < 1) {
    // Either the selection was empty, or filtering removed everything.
    const ui = SpreadsheetApp.getUi();
    ui.alert('Boffo', 'Please select cells with item barcodes.', ui.ButtonSet.OK);
    return;
  }
  log(`the user's selection contains ${numBarcodes} barcodes`);

  // Create a new sheet where results will be written.
  let resultsSheet = createResultsSheet(fields.map((field) => field[0]));
  let lastLetter = lastColumnLetter();

  // Get these values here instead of doing property lookups in the loop.
  const props = PropertiesService.getUserProperties();
  const url   = props.getProperty('boffo_folio_url');
  const id    = props.getProperty('boffo_folio_tenant_id');
  const token = props.getProperty('boffo_folio_api_token');

  // Get batches of item data for the barcodes & write to the sheet.
  const batchSize = barcodes.length > 200 ? 100 : 10;
  const barcodeBatches = batchedList(barcodes, batchSize);
  let row = 1;                          // Offset +1 for header row.
  barcodeBatches.forEach((batch, index) => {
    if (numBarcodes > 10 && batch.length >= 10) {
      note(`Getting ${nth(index + 1)} batch of ${batch.length} records` +
           ` out of a total of ${numBarcodes} …`, 10);
    } else {
      note(`Getting ${batch.length} records …`);
    }
    let records = itemRecords(batch, url, id, token);
    records.forEach(rec => {
      ++row;
      if (Object.hasOwn(rec, 'id')) {
        let cells = resultsSheet.getRange(`A${row}:${lastLetter}${row}`);
        cells.setValues([fields.map(field => field[1](rec))]);
      } else {
        let cell = resultsSheet.getRange(`A${row}`);
        cell.setValues([[rec.barcode]]);
        cell.setFontColor('red');
      }
    });
  });
  note('Done! ✨', 1);
}

/**
 * Returns the FOLIO item data for a list of barcodes.
 */
function itemRecords(barcodeList, folio_url, tenant_id, token) {
  let requests = barcodeList.map(barcode => {
    return {
      'url': `${folio_url}/inventory/items?query=barcode=${barcode}`,
      'method': 'get',
      'contentType': 'application/json',
      'headers': {
        'x-okapi-tenant': tenant_id,
        'x-okapi-token': token
      },
      'muteHttpExceptions': true
    };
  });

  log(`doing HTTP post on ${folio_url} for ${barcodeList}`);
  let responses = UrlFetchApp.fetchAll(requests);
  if (!nonempty(responses)) {
    log('fetchAll did not return any results');
    quit('Stopped due to an error',
         'The network call to FOLIO failed to return any results.' +
         ' This should not happen and probably indicates an error' +
         ' in this program. Try the operation again in case it is' +
         ' due to a transient network error; if you get this same' +
         ' error again, please report this to the developers.');
  }

  let records = [];
  responses.forEach((response, index) => {
    let barcode = barcodeList[index];

    // If there was an error, stop short and exit immediately.
    let httpCode = response.getResponseCode();
    log(`got HTTP response code ${httpCode} for ${barcode}`);
    if (httpCode >= 300) {
      note('Stopping due to error.', 0.1);
      switch (httpCode) {
        case 400:
        case 401:
        case 403:
          quit('Stopped due to a permissions error',
               'A FOLIO authentication error occurred. This can be' +
               ' due to an invalid FOLIO URL, tenant ID, or token,' +
               ' or if the user account lacks FOLIO permissions to' +
               ' perform the action requested. To fix this, try to' +
               ' reset the FOLIO credentials (using the Boffo menu' +
               ' option for that).  If this error persists, please' +
               ' contact your FOLIO administrator for assistance.');
          break;
        case 404:
          quit('Stopped due to an error',
               'The API call made by Boffo does not seem to exist at' +
               ' the address Boffo attempted to use. This may be due' +
               ' to a temporary network glitch. Please wait a moment' +
               ' then retry the same operation again. If the problem' +
               ' persists, please report this to the developers.');
          break;
        case 409:
        case 500:
        case 501:
          quit('Stopped due to a server error',
               'FOLIO turned an internal server error. This may be due' +
               ' to a temporary problem with FOLIO itself. Please wait' +
               ' a moment, then retry the same operation. If the error' +
               ' persists, please report it to the developers.' +
               ` (Error code ${httpCode}.)`);
          break;
        default:
          quit('Stopped due to an error',
               `An error occurred communicating with FOLIO` +
               ` (code ${httpCode}). Please report this to` +
               ' the developers.');
      }
    };

    let folioResult = JSON.parse(response.getContentText());
    if (folioResult.totalRecords == 0) {
      log(`Folio did not return data for ${barcode}`);
      records.push({'barcode': barcode});
    } else {
      records.push(folioResult.items[0]);
    }
  });
  return records;
}

/**
 * Creates the results sheet and returns it.
 */
function createResultsSheet(headings) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(uniqueSheetName());
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

function batchedList(input, sliceSize) {
  let output = [];
  let sliceStart = 0;
  let sliceEnd = Math.min(sliceSize, input.length);
  while (sliceStart < input.length) {
    output.push(input.slice(sliceStart, sliceEnd));
    sliceStart = sliceEnd;
    sliceEnd = sliceStart + Math.min(input.length - sliceStart, sliceSize);
  }
  return output;
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
  SpreadsheetApp.getUi().showModalDialog(dialog, 'FOLIO credentials');
}

/**
 * Returns an HTML template object for requesting FOLIO creds from the user.
 * If optional argument callAfterSuccess is not empty, it must be the name
 * of a Boffo function to be called after the credentials are successfully
 * used to create a FOLIO token.
 */
function buildCredentialsDialog(callAfterSuccess = '') {
  log(`building form for Folio creds; callAfterSuccess = ${callAfterSuccess}`);
  let htmlTemplate = HtmlService.createTemplateFromFile('credentials-form');
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
  const props = PropertiesService.getUserProperties();
  return (nonempty(props.getProperty('boffo_folio_url')) &&
          nonempty(props.getProperty('boffo_folio_tenant_id')) &&
          nonempty(props.getProperty('boffo_folio_api_token')));
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
    SpreadsheetApp.getUi().showModalDialog(dialog, 'FOLIO information needed');
  }
}

/**
 * Gets called by the submit() method in credentials-form.html. Returns true
 * if successful, false or an exception if not successful.
 *
 * IMPORTANT: this function is called from the context of the client-side
 * Javascript code in credentials-form.html. Any exceptions in this function
 * will not be caught by the spreadsheet framework; instead, they'll go to
 * the google.script.run framework's withFailureHandler() handler. This means
 * that calling ui.alert() here works (as long as the code in
 * credentials-form.html closes the dialog, which it does in the failure
 * handler) but exceptions will only show up in the logs, without the usual
 * black error banner across the top of the spreadsheet.
 */
function saveFolioInfo(url, tenantId, user, password, callAfterSuccess = '') {
  // Before saving the given info, we try to create a token. If that fails,
  // don't save the input because something is probably wrong with it.
  url = stripTrailingSlash(url);
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
  let response;
  let httpCode;
  try {
    response = UrlFetchApp.fetch(endpoint, options);
    httpCode = response.getResponseCode();
    log(`got response from Folio with HTTP code ${httpCode}`);
  } catch ({name, message}) {
    log(`error attempting UrlFetchApp on ${endpoint}: ${message}`);
    quit('Failed due to an unrecognized error', 
         'Please carefully check the values entered in the form. For example,' +
         ' is the URL correct and well-formed?  Are there any typos anywhere?' +
         ' If you cannot find the problem or it appears to be a bug in Boffo,' +
         ' please contact the developers and report the following error' +
         ` message:  ${message}`, true);
    return false;
  }

  if (httpCode < 300) {
    let response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      let token = response_headers['x-okapi-token'];
      const props = PropertiesService.getUserProperties();
      props.setProperty('boffo_folio_api_token', token);
      log('got token from Folio and saved it');
      // Also save the URL & tenant id now, since we know they work.
      props.setProperty('boffo_folio_url', url);
      props.setProperty('boffo_folio_tenant_id', tenantId);
      return true;
    } else {
      log('no token in the FOLIO response headers');
      quit('Unexpectedly failed to get a token back',
           'The call to FOLIO was successful, but FOLIO did not return' +
           ' a token. This situation should never occur and probably' +
           ' indicates a bug in Boffo. Please report this to the' +
           ' developers and describe what led to it.', true);
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
    const ui = SpreadsheetApp.getUi();
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
           'You can use the menu option "Set FOLIO credentials" to' +
           ' add valid credentials when you are ready. Until then,' +
           ' FOLIO lookup operations will fail.', true);
      return false;
    }

  } else {
    quit('Failed due to a FOLIO server or network problem',
         'This may be temporary. Try again after waiting a short time. If' +
         ' the error persists, please contact the FOLIO administrators or' +
         ' the developers of Boffo. (When reporting this error, please be' +
         ` sure to mention this was an HTTP code ${httpCode} error.)`, true);
    return false;
  }

  return haveCredentials();
}

/**
 * Runs a function given its name as a string. This is invoked by the code
 * in credentials-form.html and is needed b/c we can't pass a function object
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
  const props = PropertiesService.getUserProperties();
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
  PropertiesService.getUserProperties().deleteProperty('boffo_folio_api_token');
  const ui = SpreadsheetApp.getUi();
  ui.alert('Your stored FOLIO token has been deleted. You can use the' +
           ' menu option "Set FOLIO credentials" to generate a new one.');
}


// Menu item "About Boffo".
// ............................................................................

function menuItemShowAbout() {
  const htmlTemplate = HtmlService.createTemplateFromFile('about');
  // Setting the next variable on the template makes it available in the
  // script code embedded in the HTML source of about.html.
  htmlTemplate.boffo = getBoffoMetadata();
  const htmlContent = htmlTemplate.evaluate().setWidth(250).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlContent, 'About Boffo');
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
  return prop ? PropertiesService.getUserProperties().getProperty(prop) : "";
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
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let names = sheets.map((sheet) => sheet.getName());

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
  let errorText = 'This installation of Boffo has been damaged somehow:' +
      ' either some files are missing from the installation or one or' +
      ' more files are not in the expected format. Please report this' +
      ' error to the developers.';
  let errorThrown = new Error('Unable to continue.');

  try {  
    codemetaFile = HtmlService.createHtmlOutputFromFile('version.html');
  } catch ({name, message}) {
    log('Unable to read version.html: ' + message);
    SpreadsheetApp.getUi().alert(errorText);
    throw errorThrown;
  }
  try {
    return JSON.parse(codemetaFile.getContent());
  } catch ({name, message}) {
    log('Unable to parse JSON content of version.html: ' + message);
    SpreadsheetApp.getUi().alert(errorText);
    throw errorThrown;
  }
}

/**
 * Shows a brief note to the user. This is a separate function instead
 * of simply calling SpreadsheetApp.getActiveSpreadsheet().toast(...)
 * so that we can invoke it from credentials-form.html. The default duration
 * is 2 seconds.
 */
function note(message, duration = 2) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Boffo', duration);
    log('Displayed note to user: ' + message);
  } catch ({name, error}) {
    // I've seen "Exception: Service unavailable: Spreadsheets" errors on
    // occasion. In the present context, it's not worth doing more than
    // simply ignoring our inability to write a toast message.
    log(`Got exception ("${error}") trying to display a note to user. The` +
        ` note will not be shown. It would have been: "${message}"`);
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
  }
  Stopped.prototype = Error.prototype;
  if (showAlert) {
    SpreadsheetApp.getUi().alert(why + '. ' + details);
  }
  throw new Stopped();
}

/**
 * Logs the text message.
 */
function log(text) {
  console.log(text);
}

/**
 * Returns true if the given value is not empty, null, or undefined.
 * This works on arrays, which in JavaScript have the (IMHO) confusing
 * behvior that "x ? true : false" returns true for an empty array.
 */
function nonempty(value) {
  if (Array.isArray(value)) {
    return value.length > 0;
  } else {
    return value ? true : false;
  }
}

/**
 * Returns the given string with any trailing slash character removed
 * if the string ends in a slash; otherwise it returns the original string.
 */
function stripTrailingSlash(url) {
  return url.endsWith('/') ? url.slice(0, -1) : url;
}

/**
 * Returns a string consisting of the given number with an ordinal indicator
 * ("st", "nd", "rd", or "th") appended. This code was originally based in
 * part on a posting to Stack Overflow by user "Tomas Langkaas" on 2016-09-13
 * at https://stackoverflow.com/a/39466341/743730 
 */
function nth(n) {
  return `${n}` + (["st", "nd", "rd"][((n + 90) % 100 - 10) % 10 - 1] || "th");
}
