// @file    Code.js (shown as Code.gs in the Google Apps Script editor)
// @brief   Main file for Boffo
// @created 2023-06-08
// @license Please see the file named LICENSE in the project directory
// @website https://github.com/caltechlibrary/boffo


// FOLIO data field-handling abstractions.
// ............................................................................

const linefeed = String.fromCharCode(10);

// Internal constructor function used in the definition of "fields" below.
function Field(name, enabled, required, getValue) {
  this.name = name;
  this.enabled = enabled;
  this.required = required;
  this.getValue = getValue;
}

// Helper function to look up a subfield on a field that may not exist.
function subfield(field, subfield) {
  return field ? field[subfield] : '';
}

// Helper function to concatenate strings contained in lists of objects.
function collect(list, subfield) {
  return list ? list.map(el => el[subfield]).join('; ') : '';
}

// Helper function for the special case of notes fields.
function collectNotes(list) {
  return list ? list.map(el => el.note).join(linefeed + linefeed) : '';
}

// The next variable determines the FOLIO item record fields written into the
// results sheet. The order determines the order of the columns in the results
// sheet, and the length of this array determines the number of columns. The
// list of fields is based on inventory records for items, not storage records,
// drawn from examples in the Caltech Library FOLIO database. The values for
// the "enabled" field here are the initial defaults; users can change the
// field choices using the "Select record fields to show" menu item.
const fields = [
  //         Name                          Enabled Required getValue()
  //          ↓                                ↓      ↓       ↓
  new Field('Barcode',                        true,  true,  item => item.barcode),
  new Field('Title',                          true,  false, item => item.title),
  new Field('Call number',                    false, false, item => item.callNumber),
  new Field('Circulation notes',              false, false, item => collectNotes(item.circulationNotes)),
  new Field('Contributor names',              false, false, item => collect(item.contributorNames, 'name')),
  new Field('Discovery suppress',             false, false, item => item.discoverySuppress),
  new Field('Effective call number',          true,  false, item => item.effectiveCallNumberComponents.callNumber),
  new Field('Effective call number prefix',   false, false, item => item.effectiveCallNumberComponents.prefix),
  new Field('Effective call number suffix',   false, false, item => item.effectiveCallNumberComponents.suffix),
  new Field('Effective call number type ID',  false, false, item => item.effectiveCallNumberComponents.typeId),
  new Field('Effective location',             true,  false, item => item.effectiveLocation.name),
  new Field('Effective location ID',          false, false, item => item.effectiveLocation.id),
  new Field('Effective shelving order',       false, false, item => item.effectiveShelvingOrder),
  new Field('Electronic access',              false, false, item => collect(item.electronicAccess, 'uri')),
  new Field('Enumeration',                    true,  false, item => item.enumeration),
  new Field('Former IDs',                     false, false, item => item.formerIds.join(', ')),
  new Field('HRID',                           false, false, item => item.hrid),
  new Field('Holdings record ID',             false, false, item => item.holdingsRecordId),
  new Field('Is bound with',                  false, false, item => item.isBoundWith),
  new Field('Item level call number',         false, false, item => item.itemLevelCallNumber),
  new Field('Material type',                  true,  false, item => item.materialType.name),
  new Field('Material type ID',               false, false, item => item.materialType.id),
  new Field('Metadata: created by user ID',   false, false, item => item.metadata.createdByUserId),
  new Field('Metadata: created date',         false, false, item => item.metadata.createdDate),
  new Field('Metadata: updated by user ID',   false, false, item => item.metadata.updatedByUserId),
  new Field('Metadata: updated date',         false, false, item => item.metadata.updatedDate),
  new Field('Notes',                          false, false, item => collectNotes(item.notes)),
  new Field('Permanent loan type',            false, false, item => item.permanentLoanType.name),
  new Field('Permanent loan type ID',         false, false, item => item.permanentLoanType.id),
  new Field('Permanent location',             false, false, item => item.permanentLocation.name),
  new Field('Permanent location ID',          false, false, item => item.permanentLocation.id),
  new Field('Purchase order line identifier', false, false, item => item.purchaseOrderLineIdentifier),
  new Field('Statistical code IDs',           false, false, item => item.statisticalCodeIds.join(', ')),
  new Field('Status',                         true,  false, item => item.status.name),
  new Field('Status date',                    false, false, item => item.status.date),
  new Field('Tags',                           false, false, item => item.tags.tagList.join(', ')),
  new Field('Temporary location',             false, false, item => subfield(item.temporaryLocation, 'name')),
  new Field('Temporary location ID',          false, false, item => subfield(item.temporaryLocation, 'id')),
  new Field('UUID',                           true,  false, item => item.id),
  new Field('Year caption',                   false, false, item => item.yearCaption.join(', '))
];


// Google Sheets add-on menu definition.
// ............................................................................
// This creates the "Boffo" menu item in the Extensions menu in Google
// Sheets. Note: onInstall() gets called when the user installs the add-on
// from the Google Marketplace; onOpen() gets called when the user open a
// sheet (or do a page reload in their browser). The Google documentation at
// https://developers.google.com/apps-script/add-ons/concepts/editor-auth-lifecycle
// recommends that you call onOpen() from onInstall().

function onOpen() {
  // BEWARE: in the addItem calls below, the spaces after the icons are
  // actually 2 unbreakable spaces. This detail is invisible in the editor.
  SpreadsheetApp.getUi().createMenu('Boffo')
    .addItem('🔎 ﻿ ﻿Look up selected item barcodes', 'menuItemLookUpBarcodes')
    .addItem('🔦 ﻿ ﻿Find items in call number range', 'menuItemFindByCallNumbers')
    .addSeparator()
    .addItem('🇦︎ ﻿ ﻿Choose record fields to show', 'menuItemSelectFields')
    .addItem('🪪︎ ﻿ ﻿Set FOLIO user credentials', 'menuItemGetCredentials')
    .addItem('🧹﻿ ﻿ Clear FOLIO token', 'menuItemClearToken')
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
  const barcodes = getBarcodesFromSelection(true);
  const numBarcodes = barcodes.length;

  if (numBarcodes == 0) {
    log('nothing to do');
    return;
  }

  // Create a new sheet where results will be written.
  restoreFieldSelections();
  const enabledFields = fields.filter(f => f.enabled);
  const headings = enabledFields.map(f => f.name);
  const resultsSheet = createResultsSheet(numBarcodes, headings);
  const lastLetter = getLastColumnLetter();

  // Get these now to avoid repeatedly doing property lookups in the next loop.
  const {folioUrl, tenantId, token} = getStoredCredentials();

  // If few barcodes were fetched, it happens too fast to bother printing this.
  if (numBarcodes > 200) {
    note(`Looking up ${numBarcodes} barcodes …`);
  }

  // Each barcode lookup will consist of a CQL query of the form "barcode==N"
  // separated by "OR" with percent-encoded space characters around it. E.g.:
  //   barcode==35047019076454%20OR%20barcode==35047019076453
  // Estimating the average length of a barcode term (14-15 char barcode +
  // characters for "barcode==" and "%20OR%20") and assuming a max URL length
  // of 2048 leads to an estimate of ~65 barcodes max per query. Using the
  // number 50 is very conservative and also convenient for mental math.
  const barcodeBatches = batchedList(barcodes, 50);
  let row = 2;                          // Offset +1 for header row.
  let total = 0;
  let emptyValues = new Array(enabledFields.length - 1).fill('');
  barcodeBatches.forEach((batch, index) => {
    log(`working on rows starting with ${row}`);
    let cellValues = [];
    let records = itemRecords(batch, folioUrl, tenantId, token);
    if (records.length == 0) {
      log('received no records for this batch');
      cellValues = batch.map(barcode => [barcode, ...emptyValues]);
    } else {
      cellValues = records.map(rec => {
        // If item was found, record will have an 'id' field.
        if ('id' in rec) {
          total++;
          return enabledFields.map(f => f.getValue(rec));
        } else {
          // Not found => create a row with only the barcode in the first cell.
          return [rec.barcode, ...emptyValues];
        }
      });
    }
    let cells = resultsSheet.getRange(row, 1, cellValues.length, enabledFields.length);
    cells.setValues(cellValues);
    row += records.length;
  });

  log(`got total of ${total} records for ${numBarcodes} barcodes selected`);
  SpreadsheetApp.setActiveSheet(resultsSheet);
  if (numBarcodes > 200) {
    note(`Writing results – this may take a little longer …`);
  }
}

/**
 * Returns the FOLIO item data for a list of barcodes.
 */
function itemRecords(barcodes, folioUrl, tenantId, token) {
  let barcodeTerms = barcodes.join('%20OR%20barcode==');
  let baseUrl = `${folioUrl}/inventory/items`;
  // If there's only 1 barcode, barcodeTerms will be just that one. If > 1,
  // barcodeTerms will be "35047019076454%20OR%20barcode==35047019076453" etc
  let query = `?limit=${barcodes.length}&query=barcode==${barcodeTerms}`;
  let endpoint = baseUrl + query;
  let result = fetchJSON(endpoint, tenantId, token);
  log(`Folio returned ${result.totalRecords} records`);
  if (result.totalRecords > 0) {
    // We want the items returned in the order requested, but if a barcode
    // isn't found, it won't be in the results from Folio. So, first build a
    // temporary dictionary indexed by barcode, to be used in the next step.
    let itemsByBarcode = result.items.reduce((itemsByBarcode, item) => {
      itemsByBarcode[item.barcode] = item;
      return itemsByBarcode;
    }, {});
    // Now create a result list that is 1-1 with the list of barcodes given as
    // input, in which each element is either a full data structure (if the
    // barcode was found) or an object containing just the barcode (if not).
    return barcodes.map(barcode => {
      return barcode in itemsByBarcode ? itemsByBarcode[barcode] : {'barcode': barcode};
    });
  } else {
    return [];
  }
}

/**
 * Creates the results sheet and returns it.
 */
function createResultsSheet(numRows, headings) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(makeUniqueSheetName());

  const numColumnsNeeded = getNumColumnsForSelectedFields();
  if (sheet.getMaxColumns() < numColumnsNeeded) {
    sheet.insertColumns(1, numColumnsNeeded - sheet.getMaxColumns());
  }
  sheet.setColumnWidths(1, numColumnsNeeded, 150);
  sheet.setFrozenRows(1);

  const cells = sheet.getRange(`A1:A${numRows + 1}`);
  cells.setHorizontalAlignment('left');

  const lastLetter = getLastColumnLetter();
  const headerRow = sheet.getRange(`A1:${lastLetter}1`);
  headerRow.setValues([headings]);
  headerRow.setFontSize(10);
  headerRow.setFontColor('white');
  headerRow.setFontWeight('bold');
  headerRow.setBackground('#999999');

  return sheet;
}

/**
 * Takes an original list and returns a list of lists, where each sublist
 * contains sliceSize number of items from the original list.
 */
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


// Menu item "Find items in call number range"
// ............................................................................

/**
 * Ensures that the FOLIO credentials are set and then calls
 * findByCallNumbers().
 */
function menuItemFindByCallNumbers() {
  withCredentials(findByCallNumbers);
}

/**
 * Asks the user for a range of call numbers and a location, gets all the
 * items in that call number range at that location, and writes the results
 * to a new sheet in the spreadsheet.
 */
function findByCallNumbers(firstCN = undefined, lastCN = undefined) {
  const htmlTemplate = HtmlService.createTemplateFromFile('call-numbers-form');
  const locationsList = getLocationsList();
  // Create the body of the <select> element on the page. This consists of
  // <option> elements, one for each Folio location name.
  htmlTemplate.locationSelectorsList = locationsList.map(el => {
    return `<option value="${el.id}" class="wide">${el.name}</option>`;
  }).join('');
  const htmlContent = htmlTemplate.evaluate().setWidth(470).setHeight(290);
  log('showing dialog to get call number range');
  SpreadsheetApp.getUi().showModalDialog(htmlContent, 'Call number range');
}

/**
 * Gets a list of locations from this FOLIO instance. The value returned is
 * a sorted list of objects, where each object has the form
 *     {name: 'the name', id: 'the uuid string'}
 */
function getLocationsList() {
  const {folioUrl, tenantId, token} = getStoredCredentials();
  // 5000 is simply a high enough number to get the complete list.
  const endpoint = `${folioUrl}/locations?limit=5000`;
  const results = fetchJSON(endpoint, tenantId, token);
  if (! ('locations' in results)) {
    log('failed to get locations from FOLIO server');
    quit('Unable to get list of locations from server',
         'The request for a list of locations from the FOLIO server failed.' +
         ' This may be due to a temporary network glitch. Please wait a few' +
         ' seconds, then retry the command again. If this problem persists,' +
         ' please report it to the developers.');
  }
  const locationsList = results.locations.map(el => {
    return {name: el.name, id: el.id};
  });
  return locationsList.sort((location1, location2) => {
    return location1.name.localeCompare(location2.name);
  });
}

/**
 * Does the actual work of getting items in a call number range. This is
 * invoked from inside the HTML form "call-numbers-form.html" after getting
 * input from the user.
 */
function getItemsInCallNumberRange(firstCN, lastCN, locationId) {
  note('Looking up call numbers …');

  // First make sure the given CNs are valid.
  firstCN = getVerifiedCN(firstCN, locationId);
  lastCN  = getVerifiedCN(lastCN, locationId);

  // If we get this far, we have valid CNs. Let's saddle up and do this thing.
  const {folioUrl, tenantId, token} = getStoredCredentials();
  const baseUrl = `${folioUrl}/inventory/items`;

  function makeRangeQuery(cn1, cn2, numRecordsToGet = 100, offset = 0) {
    // 100 is the max that the Folio API will return for this query.
    return baseUrl + `?limit=${numRecordsToGet}&offset=${offset}&query=` +
      encodeURI(`effectiveLocationId==${locationId} AND ` +
                `effectiveCallNumberComponents.callNumber>="${cn1}" AND ` +
                `effectiveCallNumberComponents.callNumber<="${cn2}"` +
                ` sortBy effectiveShelvingOrder`);
  }

  // The user may have flipped the order of the CNs. Check & swap them if so.
  let endpoint = makeRangeQuery(firstCN, lastCN, 0);
  let expected = fetchJSON(endpoint, tenantId, token);
  if (expected.totalRecords > 0) {
    log(`range ${firstCN} -- ${lastCN} has ${expected.totalRecords} records`);
  } else {
    log(`swapping the order of the call numbers and trying one more time`);
    [firstCN, lastCN] = [lastCN, firstCN];
    endpoint = makeRangeQuery(firstCN, lastCN, 0);
    expected = fetchJSON(endpoint, tenantId, token);
    if (expected.totalRecords == 0) {
      // Get the location name so we can write it in the error message.
      const locName = getLocationsList().find(el => (el.id == locationId)).name;
      quit('No results for given call number range and location',
           `Searching FOLIO for the call number range ${firstCN} – ${lastCN}` +
           ` (in either order) at location "${locName}" produced no results.` +
           ' Please verify the call numbers (paying special attention to any' +
           ' space characters) as well as the location. If everything looks'  +
           ' correct, it is possible a temporary network glitch occurred; in' +
           ' that case, please wait a few seconds and try again. If the' +
           ' problem persists, please report it to the developers.');
    }
  }

  // Now we get the records for real. Do it in batches of 100.
  let records = [];
  let results;
  log(`getting ${expected.totalRecords} item records`);
  for (let offset = 0; offset < expected.totalRecords; offset += 100) {
    endpoint = makeRangeQuery(firstCN, lastCN, 100, offset);
    results = fetchJSON(endpoint, tenantId, token);
    if (results.items) {
      records = records.concat(results.items);
    } else {
      log(`unexpectedly got an empty batch during iteration`);
      quit('Failed to get complete set of records',
           `While downloading the item records for ${firstCN} – ${lastCN},` +
           ' Boffo unexpectedly received an empty batch from FOLIO. It may' +
           ' be due to a temporary network glitch, or it may be a symptom'  +
           ' of a deeper problem. Please wait a few seconds and try again.' +
           ' If the problem persists, please report it to the developers.');
    }
  }

  // Create a new sheet where results will be written.
  restoreFieldSelections();
  // Make sure Call number is shown in the results.
  setFieldEnabled('Call number', true);
  const enabledFields = fields.filter(f => f.enabled);
  const headings = enabledFields.map(f => f.name);
  const resultsSheet = createResultsSheet(records.length, headings);
  const lastLetter = getLastColumnLetter();

  log('writing results to sheet');
  let cellValues = [];
  records.forEach((record) => {
    cellValues.push(enabledFields.map(f => f.getValue(record)));
  });
  let cells = resultsSheet.getRange(2, 1, records.length, enabledFields.length);
  cells.setValues(cellValues);

  note('Done ✨');
  return true;
}

/**
 * Returns a query string to get items for a range of call numbers.
 */

/**
 * Returns the effectiveCallNumberComponents.callNumber value of an item
 * found by searching for the given call number at the location identified
 * by the given location UUID. This has the effect of verifying that an
 * item with the given call number exists in the database, and may also
 * provide a slightly more normalized version of the call number. If 
 * searching for the given call number produces more than one result, one
 * is picked at random.
 */
function getVerifiedCN(cn, locationId) {
  const items = getSampleItemsForCN(cn, locationId);
  if (items.length > 0) {
    // Found at least one item for the given CN + location. Return 1st value.
    return items[0].effectiveCallNumberComponents.callNumber;
  } else {
    // Get the location name so we can write it in the error message.
    const locName = getLocationsList().find(el => (el.id == locationId)).name;
    log(`failed to find an item with call number ${cn} at this location`);
    quit(`Could not find an item with call number ${cn} at this location`,
         'Searching in FOLIO did not return any items at the location' +
         ` "${locName}" for the call number as written. Please verify` +
         ` that "${cn}" is correct (paying special attention to space` +
         ' characters) and that the location is the correct one, then' +
         ' try the command again. If everything looks correct, it is'  +
         " possible that a temporary network glitch occurred; in that" +
         ' case, please wait a few seconds and try again. If the' +
         ' problem persists, please report it to the developers.');
    return '';
  }
}

/**
 * Searches by call number and returns a list of up to 100 item records
 * found.
 *
 * Note that this list may not be complete, if there are more than 100 items
 * with this call number at that location. The max number that can be
 * retrieved at one time via API from our Folio server is 100. We can get
 * more by using multiple API calls, but the purpose of this function is
 * to establish that the call number exists, not to get the full results.
 */
function getSampleItemsForCN(cn, locationId) {
  const {folioUrl, tenantId, token} = getStoredCredentials();

  function fetchJSONbyCN(thisCN) {
    const baseUrl = `${folioUrl}/inventory/items`;
    const query = `?limit=100&query=` +
          encodeURI(`effectiveLocationId==${locationId} AND ` +
                    `effectiveCallNumberComponents.callNumber=="${thisCN}"` +
                    ` sortBy effectiveShelvingOrder`);
    const endpoint = baseUrl + query;
    return fetchJSON(endpoint, tenantId, token);
  }

  let results = fetchJSONbyCN(cn);
  log(`got ${results.totalRecords} records`);
  if (results.totalRecords > 0) {
    return results.items;
  } else {
    // We didn't find the call number as given. This can happen for 2 reasons:
    // a) There are no items with that call number at the given location.
    // b) The c.n. is wrong somehow. E.g., if the user mistakenly split the
    //    call number in the wrong place(s), and/or the exact text of the
    //    call number in the database is itself incorrect (which can happen).
    // Approach:
    // 1) If the call number has whitespace and/or periods, we generate
    //    variants based on splitting the c.n. in different places.
    // 2) Otherwise, we have no more tricks to try for producing
    //    alternative-split versions, so we give up.
    if (/[ .]/.test(cn)) {
      log(`initial call number ${cn} not found; trying alternatives`);
      for (const candidate of makeCallNumberVariations(cn)) {
        results = fetchJSONbyCN(candidate);
        log(`searching for ${candidate} produced ${results.totalRecords} items`);
        if (results.totalRecords > 0) {
          return results.items;
        }
      }
    } else {
      quit(`Call number not found`,
           `Searching the FOLIO database for "${cn}" failed to produce a` +
           ' result. This can happen for different reasons. Perhaps that' +
           ' call number is invalid (e.g., having a space where one does' +
           ' not belong), or perhaps there are no items with that number' +
           ' at that location, or maybe a system glitch occurred. Please' +
           " check the call number carefully. If it's correct and you're" +
           ' certain it exists at that location, try to wait for a short' +
           ' time and run the command again.');
    }
  }
  return [];
}

/**
 * Returns a list of possible call numbers given an initial call number.
 * Needed because what is in the database for a given call number may
 * not be in the correct LoC form. For example, if the user enters
 *   HB171.A418 2003
 * it turns out that the correct form is
 *   HB171 .A418 2003
 * (the .A418 part is a cutter), but in our database, the item in question
 * was entered with the call number "HB171.A418 2003". The existence of
 * inconsistencies in the database means it's pointless to try to normalize
 * our user's input to a correct LoC call number form. Even if we had perfect
 * normalization of the user's input and we always sent 100% correct versions
 * in our query, the fact that the value *in the database* might be incorrect
 * means that we would never match it. So instead, the approach is to
 * generate a list of candidate variations. This function returns the list of
 * candidates, which callers try to use in an effort to find one that works.
 *
 * Examples:
 *   GV199.3                     -> GV199.3
 *
 *   GV199.F3                    -> GV199.F3
 *                                  GV199 .F3
 *
 *   HB171.A418 2003             -> HB171.A418 2003
 *                                  HB171 .A418 2003
 *
 *   GV 199.92 .B38 A3 1994      -> GV199.92 .B38 A3 1994
 *                                  GV199.92.B38 A3 1994
 *
 *   GV 199.92.F39 .F39 2011     -> GV199.92.F39 .F39 2011
 *                                  GV199.92.F39.F39 2011
 *                                  GV199.92 .F39 .F39 2011
 *                                  GV199.92 .F39.F39 2011
 *
 *   E505.5 102nd.F57 1999       -> E505.5 102nd.F57 1999
 *                                  E505.5 102nd .F57 1999
 *
 *   HB3717 1929.E37 2015        -> HB3717 1929.E37 2015
 *                                  HB3717 1929 .E37 2015
 *
 *   DT423.E26 9th.ed. 2012      -> DT423.E26 9th.ed. 2012
 *                                  DT423.E26 9th .ed. 2012
 */
function makeCallNumberVariations(givenCN) {
  let cn = givenCN;

  // Remove any space between the initial letter(s) and the first number.
  const splitClassRe = /^(?<letters>[A-Z]+)\s+(?<numbers>[0-9]+)(?<other>[^0-9])/i;
  const splitClassMatch = splitClassRe.exec(cn);
  if (splitClassMatch) {
    const matchedPortion = splitClassMatch[0];
    const restOfCN = cn.slice(matchedPortion.length - 1);
    const letters = splitClassMatch.groups.letters;
    const numbers = splitClassMatch.groups.numbers;
    cn = letters + numbers + restOfCN;
  }

  // Class numbers must start with 1-3 letters followed by 1+ digits.
  const classNumberRe = /^[a-z]{1,3}[0-9]+(\.[0-9]+)?/i;
  const classNumberMatch = classNumberRe.exec(cn);
  if (! classNumberMatch) {
    quit(`Invalid call number`,
         'Call numbers are expected to begin with 1-3 letters followed by' +
         ` at least one digit, but the given value "${givenCN}" does not.`);
  }

  // FIXME this does not produce all possible combinations. Need to revise
  // the code below, probably to use a recursive approach.

  // Look at what follows the class number. If there are substrings that have
  // the form of a dot followed by a letter, use that to create variations.
  let candidates = [];
  let match;
  const dotLetterRe = /\.[A-Z][0-9]+[^a-z0-9]/ig;
  while ((match = dotLetterRe.exec(cn))) {
    let index = match.index;
    let front = match.input.slice(0, index).trim();
    let tail  = match.input.slice(index).trim();
    candidates.push(front + tail);
    candidates.push(front + ' ' + tail);
  }
  log(`returning [${candidates.join(", ")}]`);
  return candidates;
}


// Menu item "Select record fields".
// ............................................................................

/**
 * Shows a dialog to let the user select the record fields shown in the results.
 */
function menuItemSelectFields() {
  restoreFieldSelections();
  const htmlTemplate = HtmlService.createTemplateFromFile('fields-form');
  const checkboxesList = fields.map((f, i) => 
    '<input type="checkbox" name="selections"' +
      ` value=${i}` +
      ((f.enabled || f.required) ? ' checked' : '') +
      (f.required ? ' readonly' : '') +
      `>${f.name}` +
      (f.required ? ' <em>(required)</em>' : '') +
      '<br>');
  // Setting the next variable on the template makes it available in the
  // script code embedded in the HTML source of fields-form.html.
  htmlTemplate.checkboxes = checkboxesList.join('');
  const htmlContent = htmlTemplate.evaluate().setWidth(350).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(htmlContent, 'Select fields');
}

/**
 * Saves the enabled/disabled state of fields from user properties.
 *
 * The user's record data field selections need to be persisted. The GAS
 * properties service only stores strings, which complicates storing the
 * "fields" array. This stores a JSON version of the array. Code that uses
 * this is careful never to try to use the retrived array directly, because
 * the methods in Field objects won't be preserved by the JSONification.
 * Only the "enabled" flag values are needed anyway, so it's okay.
 */
function saveFieldSelections(selections) {
  log(`saving field settings: ${selections}`);
  fields.forEach((field, index) => fields[index].enabled = selections[index]);
  const props = PropertiesService.getUserProperties();
  props.setProperty('boffo_fields', JSON.stringify(fields));
  return true;
}

/**
 * Restores the enabled/disabled state of fields from user properties.
 */
function restoreFieldSelections() {
  const props = PropertiesService.getUserProperties();
  if (props.getProperty('boffo_fields')) {
    log('found field selections in user properties -- restoring them');
    const originalFieldStates = fields.map(field => field.enabled);
    const storedFields = JSON.parse(props.getProperty('boffo_fields'));
    fields.forEach((field, index) => {
      // Sanity check in case different versions of Boffo change the fields.
      // Only restore those whose names match. If there are mismatches, only
      // some of the user's choices will get restored -- it's better than none.
      if (field.name == storedFields[index].name) {
        field.enabled = storedFields[index].enabled;
      }
    });
  }
}

/**
 * Sets a field's enabled/disabled status explicitly.
 */
function setFieldEnabled(name, status) {
  fields.forEach((field, index) => {
    if (field.name == name) {
      field.enabled = status;
      return;
    }
  });
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
  const dialog = buildCredentialsDialog('', 'FOLIO credentials');
  log('showing dialog to get credentials');
  SpreadsheetApp.getUi().showModalDialog(dialog, 'FOLIO credentials');
}

/**
 * Returns an HTML template object for requesting FOLIO creds from the user.
 * If optional argument callAfterSuccess is not empty, it must be the name
 * of a Boffo function to be called after the credentials are successfully
 * used to create a FOLIO token.
 */
function buildCredentialsDialog(callAfterSuccess, title) {
  log(`building form for Folio creds; callAfterSuccess = ${callAfterSuccess}`);
  const htmlTemplate = HtmlService.createTemplateFromFile('credentials-form');
  htmlTemplate.callAfterSuccess = callAfterSuccess;
  htmlTemplate.title = title;
  return htmlTemplate.evaluate().setWidth(475).setHeight(350);
}

/**
 * Returns multiple values needed by our calls to FOLIO. The values returned
 * are named folioUrl, tenantId, and token.
 */
function getStoredCredentials() {
  const props = PropertiesService.getUserProperties();
  return {folioUrl: props.getProperty('boffo_folio_url'),
          tenantId: props.getProperty('boffo_folio_tenant_id'),
          token:    props.getProperty('boffo_folio_api_token')};
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
  const {folioUrl, tenantId, token} = getStoredCredentials();
  const haveCreds = nonempty(folioUrl) && nonempty(tenantId) && nonempty(token);
  if (haveCreds && haveValidToken(folioUrl, tenantId, token)) {
    log(`have credentials; calling ${funcToCall.name}`);
    funcToCall();
  } else {
    const what   = ((folioUrl && tenantId) ? 'update' : 'obtain');
    const title  = `Boffo needs to ${what} FOLIO credentials`;
    log(`need to ${what} credentials before calling ${funcToCall.name}`);
    const dialog = buildCredentialsDialog(funcToCall.name, title);
    SpreadsheetApp.getUi().showModalDialog(dialog, title);
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
  const endpoint = url + '/authn/login';
  const payload = JSON.stringify({
      'tenant': tenantId,
      'username': user,
      'password': password
  });
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': payload,
    'headers': {
      'x-okapi-tenant': tenantId
    },
    'muteHttpExceptions': true
  };
  let response;
  let httpCode;

  try {
    log(`doing HTTP post on ${endpoint}`);
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
         ` message:  ${message}`);
    return false;
  }

  if (httpCode < 300) {
    const response_headers = response.getHeaders();
    if ('x-okapi-token' in response_headers) {
      const token = response_headers['x-okapi-token'];
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
           'The call to FOLIO was successful but FOLIO did not return' +
           ' a token. This should never occur, and probably indicates' +
           ' a bug in Boffo. Please report this to the developers and' +
           ' describe what led to it.');
      return false;
    }

  } else if (httpCode < 500) {
    const responseContent = response.getContentText();
    let folioMsg = responseContent;
    if (nonempty(responseContent) && responseContent.startsWith('{')) {
      const results = JSON.parse(response.getContentText());
      folioMsg = results.errors[0].message;
    }
    const question = `FOLIO rejected the request: ${folioMsg}. Try again?`;
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
           ' FOLIO lookup operations will fail.');
      return false;
    }

  } else {
    quit('Failed due to a FOLIO server or network problem',
         'This may be temporary. Try again after waiting a short time. If' +
         ' the error persists, please contact the FOLIO administrators or' +
         ' the developers of Boffo. (When reporting this error, please be' +
         ` sure to mention this was an HTTP code ${httpCode} error.)`);
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
 */
function haveValidToken(folioUrl, tenantId, token) {
  // The only way to check the token is to try to make an API call.
  const endpoint = folioUrl + '/instance-statuses?limit=0';
  const options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': tenantId,
      'x-okapi-token': token
    },
    'muteHttpExceptions': true
  };
  
  log(`testing if token is valid by doing HTTP get on ${endpoint}`);
  const response = UrlFetchApp.fetch(endpoint, options);
  const httpCode = response.getResponseCode();
  log(`got code ${httpCode}; token ${httpCode < 400 ? "is" : "is not"} valid`);
  return httpCode < 400;
}

/**
 * Returns true if it looks like the necessary data to use the FOLIO API have
 * been stored. This does NOT verify that the credentials actually work;
 * doing so requires an API call, which takes time and slows down user
 * interaction.
 */
function haveCredentials() {
  const props = PropertiesService.getUserProperties();
  return (nonempty(props.getProperty('boffo_folio_url')) &&
          nonempty(props.getProperty('boffo_folio_tenant_id')) &&
          nonempty(props.getProperty('boffo_folio_api_token')));
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
  const htmlContent = htmlTemplate.evaluate().setWidth(250).setHeight(220);
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
 * Does an HTTP call, checks for errors, and either quits or returns the
 * response.
 */
function fetchJSON(endpoint, tenantId, token) {
  const options = {
    'method': 'get',
    'contentType': 'application/json',
    'headers': {
      'x-okapi-tenant': tenantId,
      'x-okapi-token': token
    },
    'escaping': false,
    'muteHttpExceptions': true
  };
  let response;
  let httpCode;

  try {
    log(`doing HTTP ${options.method} on ${endpoint}`);
    response = UrlFetchApp.fetch(endpoint, options);
    httpCode = response.getResponseCode();
    log(`got HTTP response code ${httpCode}`);
  } catch ({name, message}) {
    log(`error attempting UrlFetchApp on ${endpoint}: ${message}`);
    quit('Failed due to an unrecognized error', 
         'Please carefully check the values entered in the form. If' +
         ' you cannot find the problem or it appears to be a bug in' +
         ' Boffo, please contact the developer and mention that you' +
         ` received the following error message:  ${message}`);
    return {};
  }

  if (httpCode >= 300) {
    note('Stopping due to error.', 0.1);
    switch (httpCode) {
    case 400:
      quit('A bug in Boffo resulted in a malformed API call to FOLIO.');
      break;
    case 401:
    case 403:
      quit('Stopped due to an authentication error',
           'The FOLIO server rejected the request. This can be' +
           ' due to an invalid token, FOLIO URL, or tenant ID,' +
           ' or if the user account lacks FOLIO permissions to' +
           ' perform the action requested. To fix this, try to' +
           ' set the FOLIO user credentials via the Boffo menu' +
           ' option for that.  If this error persists, please'  +
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
      quit('Stopped due to a FOLIO server error',
           'This may be due to a temporary network problem or a' +
           ' temporary problem with the FOLIO server. Please wait' +
           ' for a few seconds, then retry the same operation. If' +
           ' the error persists, please report it to the developers.' +
           ` (Error code ${httpCode}.)`);
      break;
    default:
      quit('Stopped due to an error',
           `An error occurred communicating with FOLIO` +
           ` (code ${httpCode}). Please report this to` +
           ' the developers.');
    }
    return {};
  }

  return JSON.parse(response.getContentText());
}

/**
 * Returns an array of barcodes based on the user's selection from the sheet.
 */
function getBarcodesFromSelection(required = false) {
  // Get array of values from spreadsheet. The list we get back is a list of
  // single-item lists, so we also flatten it.
  const selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
  let barcodes = selection.getActiveRange().getDisplayValues().flat();

  // Filter out strings that don't look like barcodes. Due to the range of
  // values that could count as a barcode, the rule is to simply ignore
  // any cells that don't contain at least one number.
  barcodes = barcodes.filter(x => /\d/.test(x));
  if (barcodes.length < 1 && required) {
    // Either the selection was empty, or filtering removed everything.
    const ui = SpreadsheetApp.getUi();
    ui.alert('Boffo', 'Please select cells containing item barcodes.',
             ui.ButtonSet.OK);
  }
  log(`the user's selection contains ${barcodes.length} barcodes`);
  return barcodes;
}

/**
 * Returns the number of columns needed to hold the fields fetched from FOLIO.
 */
function getNumColumnsForSelectedFields() {
  return fields.filter(f => f.enabled).length;
}

/**
 * Returns the spreadsheet column letter corresponding to the last column.
 */
function getLastColumnLetter() {
  const lastColIndex = getNumColumnsForSelectedFields() - 1;
  if (lastColIndex >= 26) {
    return 'A' + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(lastColIndex - 26);
  } else {
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.charAt(lastColIndex);
  }
}

/**
 * Returns a unique sheet name. The name is generated by taking a base name
 * and, if necessary, appending an integer that is incremented until the name
 * is unique.
 */
function makeUniqueSheetName(baseName = 'Item Data') {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const names = sheets.map(sheet => sheet.getName());

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
  // Ideally, we'd simply read the codemeta.json file. Unfortunately, Google
  // Apps Scripts only provides a way to read HTML files in the local script
  // directory, not JSON files. That won't stop us, though! If we add a symlink
  // in the repository named "codemeta-symlink.html" pointing to codemeta.json,
  // voilà, we can read it using HtmlService and parse the content as JSON.

  let codemetaFile = {};
  const errorText = 'This installation of Boffo has been damaged somehow:' +
      ' either some files are missing from the installation or one or' +
      ' more files are not in the expected format. Please report this' +
      ' error to the developers.';
  const errorThrown = new Error('Unable to continue.');

  try {  
    codemetaFile = HtmlService.createHtmlOutputFromFile('codemeta-symlink.html');
  } catch ({name, message}) {
    log('Unable to read codemeta-symlink.html: ' + message);
    SpreadsheetApp.getUi().alert(errorText);
    throw errorThrown;
  }
  try {
    return JSON.parse(codemetaFile.getContent());
  } catch ({name, message}) {
    log('Unable to parse JSON content of codemeta-symlink.html: ' + message);
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
    log('displayed note to user: ' + message);
  } catch ({name, error}) {
    // I've seen "Exception: Service unavailable: Spreadsheets" errors on
    // occasion. In the present context, it's not worth doing more than
    // simply ignoring our inability to write a toast message.
    log(`got exception ("${error}") trying to display a note to user. The` +
        ` note will not be shown. It would have been: "${message}"`);
  }
}

/**
 * Quits execution by throwing an Error, which causes Google to display
 * a black error banner across the top of the page.
 */
function quit(why, details = '', showAlert = true) {
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

/**
 * Returns true. This is a noop function used in some GAS 'run' calls
 * in our HTML files so that the success handlers will be invoked.
 */
function proceed() {
  return true;
}
