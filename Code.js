// @file    Code.js (shown as Code.gs in the Google Apps Script editor)
// @brief   Main file for Boffo
// @created 2023-06-08
// @license Please see the file named LICENSE in the project directory
// @website https://github.com/caltechlibrary/boffo


// FOLIO data field-handling abstractions.
// ............................................................................

const linefeed = String.fromCharCode(10);

// Limit imposed by Google sheets on number of cells in an empty sheet.
const maxGoogleSheetCells = 10000000;

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
    .addItem('🔦 ﻿ ﻿Find items by call number(s)', 'menuItemFindByCallNumbers')
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
    note(`Looking up ${numBarcodes} barcodes …`, 15);
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
    note(`Writing results – this may take a little longer …`, 5);
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
  // Prepend special value "Any" to the front of the list.
  locationsList.unshift({name: "—Any—", id: "Any"});
  // Create the body of the <select> element on the page. This consists of
  // <option> elements, one for each Folio location name.
  htmlTemplate.locationSelectorsList = locationsList.map(el => {
    return `<option value="${el.id}" class="wide">${el.name}</option>`;
  }).join('');
  const htmlContent = htmlTemplate.evaluate().setWidth(470).setHeight(290);
  log('showing dialog to get call number range');
  SpreadsheetApp.getUi().showModalDialog(htmlContent, 'Find by call number(s)');
}

/**
 * Gets a list of locations from this FOLIO instance. The value returned is
 * a sorted list of objects, where each object has the form
 *     {name: 'the name', id: 'the uuid string'}
 */
function getLocationsList() {
  const {folioUrl, tenantId, token} = getStoredCredentials();
  // Some other API calls are limited to 100 items, but this one allows more.
  // 5000 is simply a high enough number to get the complete list (I hope!).
  const endpoint = `${folioUrl}/locations?limit=5000`;
  const results = fetchJSON(endpoint, tenantId, token);
  if (! ('locations' in results)) {
    quit('Unable to get list of locations from server',
         'The request for a list of locations from FOLIO failed to return' +
         ' a result. This may be due to a sudden network glitch or other'  +
         ' failure, or it may indicate a deeper issue. Please wait a few'  +
         ' seconds, then repeat the same command. If this error repeats,' +
         ' please report it to the developers.');
  }
  const locationsList = results.locations.map(location => {
    return {name: location.name, id: location.id};
  });
  return locationsList.sort((location1, location2) => {
    return location1.name.localeCompare(location2.name);
  });
}

/**
 * Dispatches to other functions based on the current situation. This is
 * invoked from inside the HTML form "call-numbers-form.html" after getting
 * input from the user.
 */
function getItemsInCallNumberRange(firstCN, lastCN, locationId) {
  note('Searching FOLIO …', 30);
  if (!lastCN || (firstCN === lastCN)) {
    log(`doing single call number search: "${firstCN}"`);
    showItemsForCallNumber(firstCN, locationId);
  } else {
    log(`doing a range search: "${firstCN}" -> "${lastCN}"`);
    showItemsForCallNumberRange(firstCN, lastCN, locationId);
  }
}

/**
 * Gets items for a single call number.
 */
function showItemsForCallNumber(cn, locationId) {
  writeResultsSheet(getItemsForCN(cn, locationId));
}

/**
 * Gets items for a call number range.
 */
function showItemsForCallNumberRange(firstCN, lastCN, locationId) {
  // A given c.n. may return multiple items. Start by getting sorted lists
  // of items for each of the two call numbers given by the user. Note the
  // need to reverse the order in the 2nd list; this is because, if we have
  // CN1 and CN2 producing lists [CN1item1, CN1item2, CN1item3] and [CN2item1,
  // CN2item2, CN2item3], we want the range to be CN1item1 -> CN2item3.
  let firstItemList = getItemsForCN(firstCN, locationId);
  let lastItemList  = getItemsForCN(lastCN, locationId).reverse();

  // If we get this far, we have valid CNs. We use the effectiveShelvingOrder
  // to help figure out the boundaries of the range, because that's the only
  // version of the call number that will work for '>=' and '<=' searches in
  // FOLIO. However, there's a complication: in some records, the field is
  // empty. If this happens, we look in the item lists for other items with
  // the same call number and use the first with an effectiveShelvingOrder.
  let firstESO = findItemWithEffectiveShelvingOrder(firstItemList, firstCN);
  let lastESO  = findItemWithEffectiveShelvingOrder(lastItemList, lastCN);

  // If we get this far, we have values for effectiveShelvingOrder for both
  // the first and last endpoints of the search. We can proceed.
  const {folioUrl, tenantId, token} = getStoredCredentials();
  const baseUrl = `${folioUrl}/inventory/items`;
  const needLocation = (locationId == 'Any' ? false : true);

  function makeRangeQuery(eso1, eso2, limit = 0, offset = 0) {
    return baseUrl + `?limit=${limit}&offset=${offset}&query=` +
      encodeURI((needLocation ? `effectiveLocationId==${locationId} AND ` : '') +
                `effectiveShelvingOrder>="${eso1}"` +
                ` AND effectiveShelvingOrder<="${eso2}"`);
  }

  // Try the range query, and beware that the user may have swapped the c.n.'s.
  let endpoint = makeRangeQuery(firstESO, lastESO);
  let expected = fetchJSON(endpoint, tenantId, token);
  if (expected.totalRecords > 0) {
    // Success.
    log(`"${firstESO}" -> "${lastESO}" has ${expected.totalRecords} records`);
  } else {
    log(`swapping the order of the call numbers and trying one more time`);
    // Be careful about the fact that we reversed the order of lastItemList.
    [firstItemList, lastItemList] = [lastItemList.reverse(), firstItemList.reverse()];
    firstESO = findItemWithEffectiveShelvingOrder(firstItemList, lastCN);
    lastESO  = findItemWithEffectiveShelvingOrder(lastItemList, firstCN);
    endpoint = makeRangeQuery(firstESO, lastESO);
    expected = fetchJSON(endpoint, tenantId, token);
    if (expected.totalRecords > 0) {
      log(`"${firstESO}" -> "${lastESO}" has ${expected.totalRecords} records`);
    } else {
      // Get the location name so we can write it in the error message.
      let where = '';
      let alsoWhere = '';
      if (needLocation) {
        const locName = getLocationsList().find(el => (el.id == locationId)).name;
        where = ` at location "${locName}"`;
        alsoWhere = 'as well as the location';
      }
      quit('Could not find any items for this call number range',
           `Searching FOLIO for the call number range ${firstCN} – ${lastCN}`  +
           ` (in either order)${where} produced no results. Please verify the` +
           ' the call numbers (paying special attention to period and space'   +
           ' characters)${alsoWhere}. If they are all correct, it is possible' +
           ' the failure occurred due to a temporary network glitch or other'  +
           ' temporary problem. Please wait a short time, then try the same'   +
           ' search again. If this situation repeats, please report it to the' +
           ' developers.');
    }
  }

  // Don't start downloading results if we won't be able to write them.
  const numColumns = getEnabledFields().length;
  const maxRows = Math.trunc(maxGoogleSheetCells/numColumns) - 1;
  if (expected.totalRecords > maxRows) {
    quit('This query exceeds the maximum number of results that can be written',
         `The number of records this produced (${maxRows.toLocaleString()})`   +
         ` times the number of currently-selected data fields (${numColumns})` +
         ` exceeds the number of cells that a Google spreadsheet can contain.`);
  }

  // Now get the records.
  note(`Fetching ${expected.totalRecords.toLocaleString()} records from FOLIO …`, 30);
  const makeQuery = makeRangeQuery.bind(null, firstESO, lastESO);
  const records = getRecordsForQuery(makeQuery, expected.totalRecords, tenantId, token);

  // And we're done.
  writeResultsSheet(sortByShelvingOrder(records));
}

/**
 * Returns all item records for a given call number at a given location.
 */
function getItemsForCN(cn, locationId) {
  // Remember this in case we have to print a message to the user.
  const givenCN = cn;

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

  // Insert wildcards in strategic places. This is an effort to account for the
  // following common errors seen in user inputs:
  //   - missing a period (e.g., QC921.5 B9 versus QC921.5 .B9)
  //   - adding or missing spaces (e.g., GV199 .F3 versus GV199.F3)
  // Keep in mind that it doesn't matter whether the user's input follows
  // correct LoC rules, because the entries in the database may have errors
  // too. What we want is to account for errors in either source.
  cn = cn.replace(/\s+\.(?! )/g, ' *');
  cn = cn.replace(/\s+(?=[a-z])/gi, ' *');
  cn = cn.replace(/(?<=[0-9])\.(?=[a-z])/gi, '*.');

  const {folioUrl, tenantId, token} = getStoredCredentials();
  const baseUrl = `${folioUrl}/inventory/items`;
  const needLocation = (locationId == 'Any' ? false : true);

  function makeQuery(limit = 0, offset = 0) {
    // Search on the wildcarded call number at the given location.
    // 100 is the max that the Folio API will return for this query.
    return baseUrl + `?limit=${limit}&offset=${offset}&query=` +
      encodeURI((needLocation ? `effectiveLocationId==${locationId} AND ` : '') +
                `effectiveCallNumberComponents.callNumber=="${cn}"`);
  }

  // Do preliminary query to get the number of records.
  let endpoint = makeQuery();
  let results = fetchJSON(endpoint, tenantId, token);
  log(`FOLIO has ${results.totalRecords} items for ${cn} at location ${locationId}`);

  // Now get the records.
  if (results.totalRecords > 0) {
    return getRecordsForQuery(makeQuery, results.totalRecords, tenantId, token);
  } else {
    // Get the location name so we can write it in the error message.
    let where = '';
    if (needLocation) {
      const locName = getLocationsList().find(el => (el.id == locationId)).name;
      where = ` at location "${locName}"`;
    }
    quit(`Could not find an item with call number "${givenCN}"${where}`,
         'FOLIO did not return any items for the call number as written.'   +
         ` Please verify the call number "${givenCN}" (paying attention to` +
         ' periods and space characters) and the location. If they are all' +
         ' correct, it is possible the failure occurred due to a temporary' +
         ' network glitch or other temporary problem. In that case, please' +
         ' wait a short time, then try the same search again. If the error' +
         ' repeats, please report it to the developers.');
    // This branch will never actually return, but do this for consistency:
    return [];
  }
}

/**
 * Takes a function object, a total numbef of records to get, and the
 * tenant ID and token, then iterates to get all the Folio records, and
 * finally returns the array of record objects.
 *
 * The first argument (makeQuery) must be a function object that takes
 * two parameters: the "limit" value to a Folio call, and the "offset"
 * value.
 */
function getRecordsForQuery(makeQuery, totalRecords, tenantId, token) {
  let records = [];
  let results;
  for (let offset = 0; offset <= totalRecords; offset += 100) {
    results = fetchJSON(makeQuery(100, offset), tenantId, token);
    if (results.items) {
      records.push.apply(records, results.items);
    } else {
      quit(`Failed to get complete set of records`,
           ' Boffo unexpectedly received an empty batch from FOLIO. It' +
           ' may be due to a sudden network glitch or other temporary'  +
           ' failure, or it may indicate a deeper problem. Please wait' +
           ' a few seconds, then try again. If this situation repeats,' +
           ' please report it to the developers.');
    }
    if (offset > 0 && (offset % 5000) == 0) {
      note(`Fetched ${offset.toLocaleString()} records so far and still going …`, 30);
    }
  }
  return records;
}

/**
 * Takes a list of items (assumed to be sorted by effectiveShelvingOrder) and
 * returns the first value of effectiveShelvingOrder found.
 *
 * We use the effectiveShelvingOrder for things like searching by call number
 * ranges, because it's the only version of the call number that will work
 * for '>=' and '<=' searches in FOLIO. However, there's a complication: in
 * some item records, the field is empty. If that happens, our fallback
 * approach is to look in the lists of items for the next one with a nonempty
 * effectiveShelvingOrder. Now, testing the first item and only looking for a
 * fallback if it has an empty effectiveShelvingOrder field is equivalent to
 * simply iterating over the whole list looking for the first one with a
 * value. That's what this function does.
 */
function findItemWithEffectiveShelvingOrder(itemList, cn) {
  for (let i = 0; i < itemList.length; i++) {
    if (itemList[i].effectiveShelvingOrder) {
      return itemList[i].effectiveShelvingOrder;
    }
  }
  quit(`Cannot search using call number "${cn}"`,
       `None of the item records associated with the call number "${cn}"` +
       ' have a value for the record field "effectiveShelvingOrder".' +
       ' Boffo needs to use this field when searching ranges of call' +
       " numbers and can't proceed if all of the records for a given" +
       ' call number lack a value for this field.');
  return '';
}

/**
 * Sorts an array of item records by the effectiveShelvingOrder field.
 *
 * In principle, you can add "sortBy effectiveShelvingOrder" to the API
 * call to Folio. In practice, I got bizarre results when I tried. This
 * way, we know exactly what we're doing.
 */
function sortByShelvingOrder(records) {
  // sort() changes the array in place, but returning the value leads to
  // clearer calling code.
  return records.sort((r1, r2) => {
    return r1.effectiveShelvingOrder.localeCompare(r2.effectiveShelvingOrder);
  });
}

/**
 * Returns the name for a location given a location id.
 */
function getNameForLocation(locationId) {
  return getLocationsList().find(el => (el.id == locationId)).name;
}

/**
 * Writes the results of call number searches to a new sheet.
 */
function writeResultsSheet(records) {
  const enabledFields = getEnabledFields();
  const headings = enabledFields.map(f => f.name);
  let cellValues = [];
  records.forEach((record) => {
    cellValues.push(enabledFields.map(f => f.getValue(record)));
  });
  const resultsSheet = createResultsSheet(records.length, headings);
  const cells = resultsSheet.getRange(2, 1, records.length, enabledFields.length);
  cells.setValues(cellValues);
  note('Writing results to sheet ✨', 10);
  SpreadsheetApp.setActiveSheet(resultsSheet);
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
 * Return the subset of fields that is enabled, making sure to always
 * enable the Barcode field.
 */
function getEnabledFields() {
  restoreFieldSelections();
  setFieldEnabled('Barcode', true);
  return fields.filter(f => f.enabled);
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
  const haveCreds = isNonempty(folioUrl) && isNonempty(tenantId) && isNonempty(token);
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
    if (isNonempty(responseContent) && responseContent.startsWith('{')) {
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
  return (isNonempty(props.getProperty('boffo_folio_url')) &&
          isNonempty(props.getProperty('boffo_folio_tenant_id')) &&
          isNonempty(props.getProperty('boffo_folio_api_token')));
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

/**
 * Shows the About dialog for Boffo.
 */
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
    log('displaying note to user: ' + message);
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Boffo', duration);
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
  note('Boffo stopped because of an error.');
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
function isNonempty(value) {
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
