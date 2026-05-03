/////////////////////////////////////
// APPEND REVIEWED PDF TO INGREDIENTS MASTER
// FINAL PIPELINE: Extracted Lines -> Master
//
// INGREDIENTS MASTER HEADERS:
// Ingredient ID | Item Code | Ingredient | Clean Name | Category | Product Group |
// Supplier | Pack Size | Pack Qty | Pack Price (£) | Base Unit |
// Cost per Unit (£) | Unit Per Pack/Case | Notes
//
// IMPORTANT MAPPING:
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
    ui.alert('Cannot append.\n\nReview still has Pending or Needs Cloud Fix rows.');
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

    const cases = getValueByHeader_(row, extractedHeaders, 'Cases');
    const packPriceRaw = getValueByHeader_(row, extractedHeaders, 'Unit Price');
    const lineTotal = getValueByHeader_(row, extractedHeaders, 'Line Total');

    const packPrice = toPdfNumber_(packPriceRaw);

    if (!description || !supplier || !packSize || !packPrice) {
      skipped++;
      return;
    }

    /////////////////////////////////////
    // PACK SIZE PARSING
    /////////////////////////////////////

    const parsed = parsePackSizeToUnits_(packSize);

    const baseUnitRaw = getValueByHeader_(row, extractedHeaders, 'Base Unit');
    const baseUnit = parsed.baseUnit || baseUnitRaw || '';

    const packQty = parsed.packQty || '';
    const unitPerCase = parsed.unitPerCase || '';

    let costPerUnit = '';

    if (packPrice && unitPerCase) {
      costPerUnit = packPrice / Number(unitPerCase);
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
      packQty: packQty,
      packPrice: packPrice,
      baseUnit: baseUnit,
      costPerUnit: costPerUnit,
      unitPerCase: unitPerCase,
      cases: cases,
      lineTotal: lineTotal,
      packParseFlag: parsed.reviewFlag,
      packParseNotes: parsed.notes
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

  formatIngredientsMasterCostingColumns_();

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
    'Ingredient ID': makePdfIngredientId_(),
    'Item Code': data.itemCode,
    'Ingredient': data.description,
    'Clean Name': data.cleanName,
    'Supplier': data.supplier,
    'Pack Size': data.packSize,
    'Pack Qty': data.packQty,
    'Pack Price (£)': data.packPrice,
    'Base Unit': data.baseUnit,
    'Cost per Unit (£)': data.costPerUnit,
    'Unit Per Pack/Case': data.unitPerCase,
    'Notes': buildPdfAppendNote_('Imported from reviewed PDF', data)
  });

  sheet.appendRow(newRow);
}


/////////////////////////////////////
// UPDATE EXISTING INGREDIENTS MASTER ROW
/////////////////////////////////////

function updateIngredientsMasterFromPdfRow_(sheet, headers, rowNumber, data) {
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Item Code', data.itemCode);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Supplier', data.supplier);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Pack Size', data.packSize);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Pack Qty', data.packQty);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Pack Price (£)', data.packPrice);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Base Unit', data.baseUnit);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Cost per Unit (£)', data.costPerUnit);
  setCellIfHeaderExists_(sheet, headers, rowNumber, 'Unit Per Pack/Case', data.unitPerCase);

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
  let note = prefix +
    ' | Pack Price from Unit Price: £' + data.packPrice +
    ' | Cases: ' + (data.cases || '') +
    ' | Line Total: £' + (data.lineTotal || '');

  if (data.packParseFlag && data.packParseFlag !== 'OK') {
    note += ' | Pack Parse: ' + data.packParseFlag;
  }

  if (data.packParseNotes) {
    note += ' | ' + data.packParseNotes;
  }

  note += ' | ' + new Date();

  return note;
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

  sheet.getRange(rowNumber, col).setValue(noteText);
}


function makePdfIngredientId_() {
  return 'ING-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}


function toPdfNumber_(value) {
  if (value === null || value === undefined || value === '') return '';

  const cleaned = value
    .toString()
    .replace(/[£,]/g, '')
    .trim();

  const num = Number(cleaned);

  return isNaN(num) ? '' : num;
}


/////////////////////////////////////
// FORMAT INGREDIENTS MASTER COSTING COLUMNS
/////////////////////////////////////

function formatIngredientsMasterCostingColumns_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  if (!sheet) return;

  const headers = getHeaderMap_(sheet, 1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rowCount = lastRow - 1;

  if (headers['Pack Qty']) {
    sheet.getRange(2, headers['Pack Qty'], rowCount, 1).setNumberFormat('0.####');
  }

  if (headers['Pack Price (£)']) {
    sheet.getRange(2, headers['Pack Price (£)'], rowCount, 1).setNumberFormat('£0.00');
  }

  if (headers['Cost per Unit (£)']) {
    sheet.getRange(2, headers['Cost per Unit (£)'], rowCount, 1).setNumberFormat('£0.0000');
  }

  if (headers['Unit Per Pack/Case']) {
    sheet.getRange(2, headers['Unit Per Pack/Case'], rowCount, 1).setNumberFormat('0.####');
  }
}