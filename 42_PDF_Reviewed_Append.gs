/////////////////////////////////////
// APPEND REVIEWED PDF TO INGREDIENTS MASTER
// FINAL PIPELINE: Extracted Lines -> Master
//
// IMPORTANT MAPPING RULE:
// Pack Price (£) = Unit Price
// Line Total     = Line Total
// Cases          = Cases
/////////////////////////////////////

function appendReviewedPdfExtractedLinesToIngredientsMaster() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const extractedSheet = ss.getSheetByName('PDF Extracted Lines');
  const reviewSheet = ss.getSheetByName('PDF Review');
  const masterSheet = ss.getSheetByName('Ingredients Master');

  if (!extractedSheet) throw new Error('Missing sheet: PDF Extracted Lines');
  if (!reviewSheet) throw new Error('Missing sheet: PDF Review');
  if (!masterSheet) throw new Error('Missing sheet: Ingredients Master');

  const fileId = getLatestPdfExtractedDriveFileId_();

  if (!fileId) {
    ui.alert('No Drive File ID found.');
    return;
  }

  if (hasBlockingPdfReviewRows_(reviewSheet, fileId)) {
    ui.alert(
      'Cannot append.\n\nReview still has Pending or Needs Cloud Fix rows.'
    );
    return;
  }

  const response = ui.alert(
    'Append PDF to Ingredients Master?',
    'This will update and append ingredients.\n\nPack Price will be taken from Unit Price, not Line Total.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const extractedHeaders = getHeaderMap_(extractedSheet, 1);
  const masterHeaders = getHeaderMap_(masterSheet, 1);

  const lastExtractedRow = extractedSheet.getLastRow();

  if (lastExtractedRow < 2) {
    ui.alert('No rows found on PDF Extracted Lines.');
    return;
  }

  const extractedValues = extractedSheet
    .getRange(2, 1, lastExtractedRow - 1, extractedSheet.getLastColumn())
    .getValues();

  const masterValues = masterSheet.getLastRow() > 1
    ? masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues()
    : [];

  const masterLookup = buildIngredientsMasterLookup_(masterValues, masterHeaders);

  let appended = 0;
  let updated = 0;
  let ignored = 0;
  let skipped = 0;

  extractedValues.forEach(function(row) {
    const rowFileId = getValueByHeader_(row, extractedHeaders, 'Drive File ID');

    if ((rowFileId || '').toString().trim() !== fileId.toString().trim()) return;

    const reviewFlag = (getValueByHeader_(row, extractedHeaders, 'Review Flag') || '')
      .toString()
      .trim();

    if (reviewFlag === 'IGNORE') {
      ignored++;
      return;
    }

    const description = getValueByHeader_(row, extractedHeaders, 'Description');
    const supplier = getValueByHeader_(row, extractedHeaders, 'Supplier');
    const itemCode = getValueByHeader_(row, extractedHeaders, 'Item Code');
    const packSize = getValueByHeader_(row, extractedHeaders, 'Pack Size');
    const baseUnit = getValueByHeader_(row, extractedHeaders, 'Base Unit');

    /////////////////////////////////////
    // CORRECT PRICE MAPPING
    /////////////////////////////////////

    const cases = getValueByHeader_(row, extractedHeaders, 'Cases');
    const packPrice = getValueByHeader_(row, extractedHeaders, 'Unit Price');
    const lineTotal = getValueByHeader_(row, extractedHeaders, 'Line Total');

    if (!description || !supplier || !packSize || !packPrice) {
      skipped++;
      return;
    }

    const cleanName = cleanIngredientNameForPdfAppend_(description);

    const codeKey = makePdfMasterCodeKey_(supplier, itemCode);
    const fallbackKey = makePdfMasterFallbackKey_(supplier, cleanName, packSize, baseUnit);

    const existingRow =
      itemCode && masterLookup.byCode[codeKey]
        ? masterLookup.byCode[codeKey]
        : masterLookup.byFallback[fallbackKey];

    const data = {
      supplier: supplier,
      itemCode: itemCode,
      description: description,
      cleanName: cleanName,
      packSize: packSize,
      baseUnit: baseUnit,
      packPrice: packPrice,
      cases: cases,
      lineTotal: lineTotal
    };

    if (existingRow) {
      updateIngredientsMasterFromPdfRow_(
        masterSheet,
        masterHeaders,
        existingRow,
        data
      );
      updated++;
    } else {
      appendIngredientsMasterFromPdfRow_(
        masterSheet,
        masterHeaders,
        data
      );
      appended++;
    }
  });

  ui.alert(
    'Append complete\n\n' +
    'Updated: ' + updated + '\n' +
    'Appended: ' + appended + '\n' +
    'Ignored: ' + ignored + '\n' +
    'Skipped: ' + skipped + '\n\n' +
    'Pack Price mapped from Unit Price.'
  );
}


/////////////////////////////////////
// APPEND HELPERS
/////////////////////////////////////

function getLatestPdfExtractedDriveFileId_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Extracted Lines');

  if (!sheet || sheet.getLastRow() < 2) return '';

  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Extracted Lines');

  return sheet.getRange(sheet.getLastRow(), fileIdCol).getValue();
}


function hasBlockingPdfReviewRows_(reviewSheet, fileId) {
  const lastRow = reviewSheet.getLastRow();
  if (lastRow < 2) return false;

  const headers = getHeaderMap_(reviewSheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Review');
  const statusCol = getRequiredHeader_(headers, 'Review Status', 'PDF Review');

  const values = reviewSheet
    .getRange(2, 1, lastRow - 1, reviewSheet.getLastColumn())
    .getValues();

  return values.some(function(row) {
    const rowFileId = row[fileIdCol - 1];
    const status = (row[statusCol - 1] || '').toString().trim();

    return (
      (rowFileId || '').toString().trim() === fileId.toString().trim() &&
      (status === 'Pending' || status === 'Needs Cloud Fix')
    );
  });
}


/////////////////////////////////////
// BUILD INGREDIENTS MASTER LOOKUP
/////////////////////////////////////

function buildIngredientsMasterLookup_(masterValues, masterHeaders) {
  const lookup = {
    byCode: {},
    byFallback: {}
  };

  masterValues.forEach(function(row, index) {
    const sheetRow = index + 2;

    const supplier = getValueByHeader_(row, masterHeaders, 'Supplier');
    const itemCode = getOptionalValueByHeader_(row, masterHeaders, 'Item Code');
    const cleanName = getValueByHeader_(row, masterHeaders, 'Clean Name');
    const packSize = getValueByHeader_(row, masterHeaders, 'Pack Size');
    const baseUnit = getValueByHeader_(row, masterHeaders, 'Base Unit');

    const codeKey = makePdfMasterCodeKey_(supplier, itemCode);
    const fallbackKey = makePdfMasterFallbackKey_(supplier, cleanName, packSize, baseUnit);

    if (supplier && itemCode && !lookup.byCode[codeKey]) {
      lookup.byCode[codeKey] = sheetRow;
    }

    if (supplier && cleanName && packSize && !lookup.byFallback[fallbackKey]) {
      lookup.byFallback[fallbackKey] = sheetRow;
    }
  });

  return lookup;
}


/////////////////////////////////////
// MATCHING HELPERS
/////////////////////////////////////

function makePdfMasterCodeKey_(supplier, itemCode) {
  return [
    normalisePdfKeyPart_(supplier),
    normalisePdfKeyPart_(itemCode)
  ].join('|');
}


function makePdfMasterFallbackKey_(supplier, cleanName, packSize, baseUnit) {
  return [
    normalisePdfKeyPart_(supplier),
    normalisePdfKeyPart_(cleanName),
    normalisePdfKeyPart_(packSize),
    normalisePdfKeyPart_(baseUnit)
  ].join('|');
}


function normalisePdfKeyPart_(value) {
  return (value || '')
    .toString()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}


/////////////////////////////////////
// CLEAN INGREDIENT NAME FOR PDF APPEND
/////////////////////////////////////

function cleanIngredientNameForPdfAppend_(name) {
  return (name || '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\b\d+(\.\d+)?\s*(kg|g|ltr|lt|l|ml|cm|m|x|pk|pack|case|box|bag)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/////////////////////////////////////
// APPEND NEW INGREDIENTS MASTER ROW
/////////////////////////////////////

function appendIngredientsMasterFromPdfRow_(sheet, headers, data) {
  const newRow = new Array(sheet.getLastColumn()).fill('');

  setRowByHeaders_(newRow, headers, {
    'Ingredient': data.description,
    'Clean Name': data.cleanName,
    'Supplier': data.supplier,
    'Pack Size': data.packSize,
    'Pack Price (£)': data.packPrice,
    'Base Unit': data.baseUnit,
    'Item Code': data.itemCode,
    'Notes': buildPdfAppendNote_('Imported from reviewed PDF', data)
  });

  const idCol = headers['Ingredient ID'];

  if (idCol) {
    newRow[idCol - 1] = makePdfIngredientId_();
  }

  sheet.appendRow(newRow);
}


/////////////////////////////////////
// UPDATE EXISTING INGREDIENTS MASTER ROW
/////////////////////////////////////

function updateIngredientsMasterFromPdfRow_(sheet, headers, rowNumber, data) {
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Supplier', data.supplier);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Item Code', data.itemCode);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Pack Size', data.packSize);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Pack Price (£)', data.packPrice);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Base Unit', data.baseUnit);

  appendNoteIfHeaderExists_(
    sheet,
    headers,
    rowNumber,
    buildPdfAppendNote_('Updated from reviewed PDF', data)
  );
}


/////////////////////////////////////
// PDF APPEND NOTE
/////////////////////////////////////

function buildPdfAppendNote_(prefix, data) {
  return prefix +
    ' | Pack Price from Unit Price: £' + data.packPrice +
    ' | Cases: ' + (data.cases || '') +
    ' | Line Total: £' + (data.lineTotal || '') +
    ' | ' + new Date();
}


/////////////////////////////////////
// SAFE WRITE HELPERS
/////////////////////////////////////

function setCellIfHeaderExists_(sheet, headers, rowNumber, headerName, value) {
  const col = headers[headerName];
  if (!col) return;

  sheet.getRange(rowNumber, col).setValue(value);
}


function appendNoteIfHeaderExists_(sheet, headers, rowNumber, noteText) {
  const col = headers['Notes'];
  if (!col) return;

  const cell = sheet.getRange(rowNumber, col);
  const existing = (cell.getValue() || '').toString();

  const cleanExisting = existing
    .replace(/\n/g, ' | ')
    .replace(/\s*\|\s*\|\s*/g, ' | ')
    .trim();

  const newValue = cleanExisting
    ? cleanExisting + ' | ' + noteText
    : noteText;

  cell.setValue(newValue);
}


function makePdfIngredientId_() {
  return 'ING-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}