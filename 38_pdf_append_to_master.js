/////////////////////////////////////
// APPEND PDF EXTRACTED LINES TO INGREDIENTS MASTER
// NEW PIPELINE: PDF Extracted Lines -> Ingredients Master
/////////////////////////////////////

function appendPdfExtractedLinesToIngredientsMaster() {
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
    ui.alert('No Drive File ID found in PDF Extracted Lines.');
    return;
  }

  if (hasBlockingPdfReviewRows_(reviewSheet, fileId)) {
    ui.alert(
      'Cannot append yet.\n\n' +
      'PDF Review still has Pending or Needs Cloud Fix rows for this file.\n\n' +
      'Drive File ID:\n' + fileId
    );
    return;
  }

  const response = ui.alert(
    'Append PDF Rows to Ingredients Master?',
    'This will update matching ingredients and append new ones from PDF Extracted Lines.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const extractedHeaders = getHeaderMap_(extractedSheet, 1);
  const masterHeaders = getHeaderMap_(masterSheet, 1);

  validatePdfAppendHeaders_(extractedHeaders, masterHeaders);

  const extractedValues = extractedSheet
    .getRange(2, 1, extractedSheet.getLastRow() - 1, extractedSheet.getLastColumn())
    .getValues();

  const masterValues = masterSheet.getLastRow() > 1
    ? masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues()
    : [];

  const masterLookup = buildIngredientsMasterLookup_(masterValues, masterHeaders);

  let appended = 0;
  let updated = 0;
  let ignored = 0;
  let skipped = 0;

  extractedValues.forEach(row => {
    const rowFileId = getValueByHeader_(row, extractedHeaders, 'Drive File ID');

    if ((rowFileId || '').toString().trim() !== fileId.toString().trim()) return;

    const reviewFlag = (getValueByHeader_(row, extractedHeaders, 'Review Flag') || '').toString().trim();

    if (reviewFlag === 'IGNORE') {
      ignored++;
      return;
    }

    const description = getValueByHeader_(row, extractedHeaders, 'Description');
    const supplier = getValueByHeader_(row, extractedHeaders, 'Supplier');
    const itemCode = getValueByHeader_(row, extractedHeaders, 'Item Code');
    const packSize = getValueByHeader_(row, extractedHeaders, 'Pack Size');
    const baseUnit = getValueByHeader_(row, extractedHeaders, 'Base Unit');
    const packPrice = getValueByHeader_(row, extractedHeaders, 'Line Total') || getValueByHeader_(row, extractedHeaders, 'Unit Price');

    if (!description || !supplier || !packSize || !packPrice) {
      skipped++;
      return;
    }

    const cleanName = cleanIngredientNameForPdfAppend_(description);

    const codeKey = makePdfMasterCodeKey_(supplier, itemCode);
    const fallbackKey = makePdfMasterFallbackKey_(supplier, cleanName, packSize, baseUnit);

    const existingRowNumber =
      itemCode && masterLookup.byCode[codeKey]
        ? masterLookup.byCode[codeKey]
        : masterLookup.byFallback[fallbackKey];

    if (existingRowNumber) {
      updateIngredientsMasterFromPdfRow_(
        masterSheet,
        masterHeaders,
        existingRowNumber,
        {
          supplier,
          itemCode,
          description,
          cleanName,
          packSize,
          baseUnit,
          packPrice
        }
      );

      updated++;
      return;
    }

    appendIngredientsMasterFromPdfRow_(
      masterSheet,
      masterHeaders,
      {
        supplier,
        itemCode,
        description,
        cleanName,
        packSize,
        baseUnit,
        packPrice
      }
    );

    appended++;
  });

  ui.alert(
    'PDF rows processed into Ingredients Master.\n\n' +
    'Drive File ID:\n' + fileId + '\n\n' +
    'Updated: ' + updated + '\n' +
    'Appended: ' + appended + '\n' +
    'Ignored: ' + ignored + '\n' +
    'Skipped: ' + skipped
  );
}

/////////////////////////////////////
// GET LATEST PDF EXTRACTED DRIVE FILE ID
/////////////////////////////////////

function getLatestPdfExtractedDriveFileId_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Extracted Lines');

  if (!sheet || sheet.getLastRow() < 2) return '';

  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Extracted Lines');

  return sheet.getRange(sheet.getLastRow(), fileIdCol).getValue();
}

/////////////////////////////////////
// CHECK BLOCKING REVIEW ROWS
/////////////////////////////////////

function hasBlockingPdfReviewRows_(reviewSheet, fileId) {
  const lastRow = reviewSheet.getLastRow();
  if (lastRow < 2) return false;

  const headers = getHeaderMap_(reviewSheet, 1);

  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Review');
  const statusCol = getRequiredHeader_(headers, 'Review Status', 'PDF Review');

  const values = reviewSheet
    .getRange(2, 1, lastRow - 1, reviewSheet.getLastColumn())
    .getValues();

  return values.some(row => {
    const rowFileId = row[fileIdCol - 1];
    const status = (row[statusCol - 1] || '').toString().trim();

    return (
      (rowFileId || '').toString().trim() === fileId.toString().trim() &&
      (status === 'Pending' || status === 'Needs Cloud Fix')
    );
  });
}

/////////////////////////////////////
// VALIDATE HEADERS
/////////////////////////////////////

function validatePdfAppendHeaders_(extractedHeaders, masterHeaders) {
  [
    'Drive File ID',
    'Supplier',
    'Description',
    'Pack Size',
    'Item Code',
    'Line Total',
    'Unit Price',
    'Base Unit',
    'Review Flag'
  ].forEach(h => getRequiredHeader_(extractedHeaders, h, 'PDF Extracted Lines'));

  [
    'Ingredient',
    'Clean Name',
    'Supplier',
    'Pack Size',
    'Pack Price (£)',
    'Base Unit'
  ].forEach(h => getRequiredHeader_(masterHeaders, h, 'Ingredients Master'));

  // Optional but recommended
  getOptionalHeaderForPdfAppend_(masterHeaders, 'Ingredient ID');
  getOptionalHeaderForPdfAppend_(masterHeaders, 'Item Code');
  getOptionalHeaderForPdfAppend_(masterHeaders, 'Notes');
}

/////////////////////////////////////
// BUILD MASTER LOOKUP
/////////////////////////////////////

function buildIngredientsMasterLookup_(masterValues, masterHeaders) {
  const lookup = {
    byCode: {},
    byFallback: {}
  };

  masterValues.forEach((row, index) => {
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
// UPDATE EXISTING MASTER ROW
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
    'Updated from PDF Extracted Lines | ' + new Date()
  );
}

/////////////////////////////////////
// APPEND NEW MASTER ROW
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
    'Notes': 'Imported from PDF Extracted Lines | ' + new Date()
  });

  const idCol = headers['Ingredient ID'];
  if (idCol) {
    newRow[idCol - 1] = makePdfIngredientId_();
  }

  sheet.appendRow(newRow);
}

/////////////////////////////////////
// SAFE CELL HELPERS
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
  const existing = cell.getValue();

  cell.setValue(existing ? existing + '\n' + noteText : noteText);
}

function getOptionalHeaderForPdfAppend_(headers, headerName) {
  return headers[headerName] || null;
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

function cleanIngredientNameForPdfAppend_(name) {
  return (name || '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\b\d+(\.\d+)?\s*(kg|g|ltr|lt|l|ml|cm|m|x|pk|pack|case|box|bag)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function makePdfIngredientId_() {
  return 'ING-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}