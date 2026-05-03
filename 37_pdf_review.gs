/////////////////////////////////////
// BUILD PDF REVIEW FROM EXTRACTED LINES
// NOW ALSO CHECKS PACK SIZE PARSING
/////////////////////////////////////

function buildPdfReviewFromExtractedLines(fileId) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  if (!fileId) throw new Error('Missing fileId.');

  const sourceSheet = ss.getSheetByName('PDF Extracted Lines');
  const reviewSheet = ss.getSheetByName('PDF Review');

  if (!sourceSheet) throw new Error('Sheet "PDF Extracted Lines" not found.');
  if (!reviewSheet) throw new Error('Sheet "PDF Review" not found.');

  const sourceHeaders = getHeaderMap_(sourceSheet, 1);
  const reviewHeaders = getHeaderMap_(reviewSheet, 1);

  [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'Row No',
    'Cases',
    'Units / Weight',
    'Base Unit',
    'Description',
    'Pack Size',
    'Item Code',
    'Unit Price',
    'Line Total',
    'VAT',
    'VAT Total',
    'Review Flag'
  ].forEach(h => getRequiredHeader_(sourceHeaders, h, 'PDF Extracted Lines'));

  [
    'Review ID',
    'File Name',
    'Supplier',
    'Drive File ID',
    'Row No',
    'Original Cases',
    'Original Units / Weight',
    'Original Description',
    'Original Pack Size',
    'Original Item Code',
    'Original Unit Price',
    'Original Line Total',
    'Corrected Cases',
    'Corrected Units / Weight',
    'Corrected Description',
    'Corrected Pack Size',
    'Corrected Item Code',
    'Corrected Unit Price',
    'Corrected Line Total',
    'Review Status',
    'Reviewed By',
    'Reviewed Time',
    'Notes'
  ].forEach(h => getRequiredHeader_(reviewHeaders, h, 'PDF Review'));

  clearPdfReviewRowsForFile_(reviewSheet, reviewHeaders, fileId);

  const lastRow = sourceSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No extracted lines found.');
    return 0;
  }

  const values = sourceSheet
    .getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn())
    .getValues();

  const output = [];
  const popupLines = [];

  values.forEach(function(row) {
    const sourceFileId = getValueByHeader_(row, sourceHeaders, 'Drive File ID');

    if ((sourceFileId || '').toString().trim() !== fileId.toString().trim()) return;

    const rowNo = getValueByHeader_(row, sourceHeaders, 'Row No');
    const reviewFlag = (getValueByHeader_(row, sourceHeaders, 'Review Flag') || '').toString().trim();

    const cases = getValueByHeader_(row, sourceHeaders, 'Cases');
    const unitsWeight = getValueByHeader_(row, sourceHeaders, 'Units / Weight');
    const description = getValueByHeader_(row, sourceHeaders, 'Description');
    const packSize = getValueByHeader_(row, sourceHeaders, 'Pack Size');
    const itemCode = getValueByHeader_(row, sourceHeaders, 'Item Code');
    const unitPrice = getValueByHeader_(row, sourceHeaders, 'Unit Price');
    const lineTotal = getValueByHeader_(row, sourceHeaders, 'Line Total');

    const missing = [];
    const notes = [];

    if (
        reviewFlag &&
        reviewFlag !== 'OK' &&
        reviewFlag !== 'CHECK PACK SIZE'
      ) {
        notes.push(reviewFlag);
      }

    const hasCases = (cases || '').toString().trim() !== '';
    const hasUnitsWeight = (unitsWeight || '').toString().trim() !== '';

    if (!hasCases && !hasUnitsWeight) {
      missing.push('Quantity');
    }

    if (!description) missing.push('Description');
    if (!packSize) missing.push('Pack Size');
    if (!unitPrice) missing.push('Unit Price');

    /////////////////////////////////////
    // PACK SIZE PARSE CHECK
    /////////////////////////////////////

    if (packSize) {
      const parsedPack = parsePackSizeToUnits_(packSize);

      if (parsedPack.reviewFlag !== 'OK') {
        notes.push('CHECK PACK SIZE: ' + parsedPack.notes);
      }

      if (!parsedPack.unitPerCase) {
        notes.push('Missing Unit Per Pack/Case from pack size');
      }
    }

    if (missing.length) {
      notes.push('Missing: ' + missing.join(', '));
    }

    if (!notes.length) return;

    const reviewRow = new Array(reviewSheet.getLastColumn()).fill('');

    setRowByHeaders_(reviewRow, reviewHeaders, {
      'Review ID': fileId + '-' + rowNo,
      'File Name': getValueByHeader_(row, sourceHeaders, 'File Name'),
      'Supplier': getValueByHeader_(row, sourceHeaders, 'Supplier'),
      'Drive File ID': fileId,
      'Row No': rowNo,

      'Original Cases': cases,
      'Original Units / Weight': unitsWeight,
      'Original Description': description,
      'Original Pack Size': packSize,
      'Original Item Code': itemCode,
      'Original Unit Price': unitPrice,
      'Original Line Total': lineTotal,

      'Corrected Cases': cases,
      'Corrected Units / Weight': unitsWeight,
      'Corrected Description': description,
      'Corrected Pack Size': packSize,
      'Corrected Item Code': itemCode,
      'Corrected Unit Price': unitPrice,
      'Corrected Line Total': lineTotal,

      'Review Status': 'Pending',
      'Notes': notes.join(' | ')
    });

    output.push(reviewRow);

    popupLines.push(
      'Row ' + rowNo + ': ' + (description || '[No description]') + '\n' +
      notes.join(' | ')
    );
  });

  if (!output.length) {
    ui.alert('PDF Review built.\n\nNo review rows needed for this PDF.');
    return 0;
  }

  const startRow = Math.max(reviewSheet.getLastRow() + 1, 2);

  reviewSheet
    .getRange(startRow, 1, output.length, output[0].length)
    .setValues(output);

  ui.alert(
    'PDF Review built from Extracted Lines.\n\n' +
    'Rows needing review: ' + output.length + '\n\n' +
    popupLines.slice(0, 15).join('\n\n') +
    (popupLines.length > 15 ? '\n\nMore rows exist in PDF Review.' : '')
  );

  return output.length;
}

/////////////////////////////////////
// CLEAR PDF REVIEW ROWS FOR FILE
/////////////////////////////////////

function clearPdfReviewRowsForFile_(sheet, headerMap, fileId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const fileIdCol = getRequiredHeader_(headerMap, 'Drive File ID', 'PDF Review');

  const fileIds = sheet
    .getRange(2, fileIdCol, lastRow - 1, 1)
    .getValues()
    .flat();

  const rowsToDelete = [];

  fileIds.forEach((value, index) => {
    if ((value || '').toString().trim() === fileId.toString().trim()) {
      rowsToDelete.push(index + 2);
    }
  });

  if (!rowsToDelete.length) return;

  deleteRowsInGroups_(sheet, rowsToDelete);
}

/////////////////////////////////////
// RUN BUILD PDF REVIEW FROM EXTRACTED LINES
/////////////////////////////////////

function runBuildPdfReviewFromExtractedLines() {
  const fileId = Browser.inputBox('Enter Drive File ID to build PDF Review');

  if (!fileId || fileId === 'cancel') return;

  buildPdfReviewFromExtractedLines(fileId);
}



/////////////////////////////////////
// PDF REVIEW SYSTEM
/////////////////////////////////////

const PDF_REVIEW_SHEET_NAME_ = 'PDF Review';
const PDF_PARSED_ROWS_SHEET_NAME_ = 'PDF Parsed Rows';

/////////////////////////////////////
// PDF REVIEW HEADERS
/////////////////////////////////////

function getPdfReviewHeaders_() {
  return [
    'Review ID',
    'File Name',
    'Supplier',
    'Drive File ID',
    'Row No',

    'Original Cases',
    'Original Units / Weight',
    'Original Description',
    'Original Pack Size',
    'Original Item Code',
    'Original Unit Price',
    'Original Line Total',

    'Corrected Cases',
    'Corrected Units / Weight',
    'Corrected Description',
    'Corrected Pack Size',
    'Corrected Item Code',
    'Corrected Unit Price',
    'Corrected Line Total',

    'Review Status',
    'Reviewed By',
    'Reviewed Time',
    'Notes'
  ];
}

/////////////////////////////////////
// SETUP PDF REVIEW SHEET
/////////////////////////////////////

function setupPdfReviewSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(PDF_REVIEW_SHEET_NAME_);

  if (!sheet) {
    sheet = ss.insertSheet(PDF_REVIEW_SHEET_NAME_);
  }

  const headers = getPdfReviewHeaders_();

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  const statusCol = headers.indexOf('Review Status') + 1;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Ignore Row', 'Needs Cloud Fix'], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, statusCol, 50, 1).setDataValidation(rule);

  applyPdfReviewConditionalFormatting_(sheet);
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert('PDF Review sheet setup complete.');
}

/////////////////////////////////////
// BUILD PDF REVIEW SHEET
/////////////////////////////////////

function buildPdfReviewSheet() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const parsedSheet = ss.getSheetByName(PDF_PARSED_ROWS_SHEET_NAME_);
  const reviewSheet = ss.getSheetByName(PDF_REVIEW_SHEET_NAME_);

  if (!parsedSheet) throw new Error('Sheet "PDF Parsed Rows" not found.');
  if (!reviewSheet) throw new Error('Sheet "PDF Review" not found. Run setupPdfReviewSheet first.');

  const parsedHeaders = getHeaderMap_(parsedSheet, 1);
  const reviewHeaders = getHeaderMap_(reviewSheet, 1);

  const requiredParsedHeaders = [
  'File Name',
  'Supplier',
  'Drive File ID',
  'Row No',
  'Cases',
  'Units / Weight',
  'Description',
  'Pack Size',
  'Unit Price',
  'Line Total',
  'Status',
  'Notes'
]

  requiredParsedHeaders.forEach(function(h) {
    getRequiredHeader_(parsedHeaders, h, 'PDF Parsed Rows');
  });

  getPdfReviewHeaders_().forEach(function(h) {
    getRequiredHeader_(reviewHeaders, h, 'PDF Review');
  });

  const lastRow = parsedSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No parsed rows found.');
    return;
  }

  const data = parsedSheet
    .getRange(2, 1, lastRow - 1, parsedSheet.getLastColumn())
    .getValues();

  const reviewRows = [];

  data.forEach(function(row) {
    const status = (row[parsedHeaders['Status'] - 1] || '')
      .toString()
      .trim()
      .toUpperCase();

    const cleanedNotes = cleanPdfReviewNotes_(row[parsedHeaders['Notes'] - 1]);

      if (status !== 'CHECK') return;
      if (!cleanedNotes) return;

    const fileName = row[parsedHeaders['File Name'] - 1] || '';
    const supplier = row[parsedHeaders['Supplier'] - 1] || '';
    const fileId = row[parsedHeaders['Drive File ID'] - 1] || '';
    const rowNo = row[parsedHeaders['Row No'] - 1] || '';

    const reviewId = [fileId, rowNo].join('|');

    reviewRows.push([
      reviewId,
      fileName,
      supplier,
      fileId,
      rowNo,

      row[parsedHeaders['Cases'] - 1] || '',
      row[parsedHeaders['Units / Weight'] - 1] || '',
      row[parsedHeaders['Description'] - 1] || '',
      row[parsedHeaders['Pack Size'] - 1] || '',
      row[parsedHeaders['Item Code'] - 1] || '',
      row[parsedHeaders['Unit Price'] - 1] || '',
      row[parsedHeaders['Line Total'] - 1] || '',

      row[parsedHeaders['Cases'] - 1] || '',
      row[parsedHeaders['Units / Weight'] - 1] || '',
      row[parsedHeaders['Description'] - 1] || '',
      row[parsedHeaders['Pack Size'] - 1] || '',
      row[parsedHeaders['Item Code'] - 1] || '',
      row[parsedHeaders['Unit Price'] - 1] || '',
      row[parsedHeaders['Line Total'] - 1] || '',

      'Pending',
      '',
      '',
      cleanedNotes
    ]);
  });

 /////////////////////////////////////
// WRITE TO SHEET
/////////////////////////////////////

if (reviewSheet.getLastRow() > 1) {
  reviewSheet
    .getRange(2, 1, reviewSheet.getLastRow() - 1, reviewSheet.getLastColumn())
    .clearContent()
    .clearDataValidations();
}

if (reviewRows.length) {
  reviewSheet
    .getRange(2, 1, reviewRows.length, reviewRows[0].length)
    .setValues(reviewRows);

  const statusCol = reviewHeaders['Review Status'];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Ignore Row', 'Needs Cloud Fix'], true)
    .setAllowInvalid(false)
    .build();

  reviewSheet
    .getRange(2, statusCol, reviewRows.length, 1)
    .setDataValidation(rule);
}

applyPdfReviewConditionalFormatting_(reviewSheet);
reviewSheet.autoResizeColumns(1, getPdfReviewHeaders_().length);

ui.alert('PDF Review built.\n\nRows needing review: ' + reviewRows.length);
}



/////////////////////////////////////
// APPLY PDF REVIEW CORRECTIONS
// NEW PIPELINE: PDF Review -> PDF Extracted Lines
/////////////////////////////////////

function applyPdfReviewCorrections() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const reviewSheet = ss.getSheetByName('PDF Review');
  const extractedSheet = ss.getSheetByName('PDF Extracted Lines');

  if (!reviewSheet) throw new Error('Missing sheet: PDF Review');
  if (!extractedSheet) throw new Error('Missing sheet: PDF Extracted Lines');

  const reviewHeaders = getHeaderMap_(reviewSheet, 1);
  const extractedHeaders = getHeaderMap_(extractedSheet, 1);

  const reviewLastRow = reviewSheet.getLastRow();
  const extractedLastRow = extractedSheet.getLastRow();

  if (reviewLastRow < 2) {
    ui.alert('No review rows to apply.');
    return;
  }

  if (extractedLastRow < 2) {
    ui.alert('No extracted lines found.');
    return;
  }

  const response = ui.alert(
    'Apply PDF Review Corrections?',
    'This will update PDF Extracted Lines using Approved and Ignore Row review rows.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const reviewValues = reviewSheet
    .getRange(2, 1, reviewLastRow - 1, reviewSheet.getLastColumn())
    .getValues();

  const extractedValues = extractedSheet
    .getRange(2, 1, extractedLastRow - 1, extractedSheet.getLastColumn())
    .getValues();

  const extractedLookup = {};

  extractedValues.forEach((row, index) => {
    const fileId = getValueByHeader_(row, extractedHeaders, 'Drive File ID');
    const rowNo = getValueByHeader_(row, extractedHeaders, 'Row No');

    if (fileId && rowNo) {
      extractedLookup[fileId + '|' + rowNo] = index + 2;
    }
  });

  let applied = 0;
  let ignored = 0;
  let skipped = 0;

  reviewValues.forEach(reviewRow => {
    const status = (getValueByHeader_(reviewRow, reviewHeaders, 'Review Status') || '').toString().trim();

    if (status !== 'Approved' && status !== 'Ignore Row') {
      skipped++;
      return;
    }

    const fileId = getValueByHeader_(reviewRow, reviewHeaders, 'Drive File ID');
    const rowNo = getValueByHeader_(reviewRow, reviewHeaders, 'Row No');
    const targetRow = extractedLookup[fileId + '|' + rowNo];

    if (!targetRow) {
      skipped++;
      return;
    }

    if (status === 'Ignore Row') {
      extractedSheet
        .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Review Flag', 'PDF Extracted Lines'))
        .setValue('IGNORE');

      ignored++;
      return;
    }

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Cases', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Cases'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Units / Weight', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Units / Weight'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Description', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Description'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Pack Size', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Pack Size'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Item Code', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Item Code'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Unit Price', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Unit Price'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Line Total', 'PDF Extracted Lines'))
      .setValue(getValueByHeader_(reviewRow, reviewHeaders, 'Corrected Line Total'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Review Flag', 'PDF Extracted Lines'))
      .setValue('OK');

    applied++;
  });

  ui.alert(
    'PDF Review corrections applied.\n\n' +
    'Approved rows applied: ' + applied + '\n' +
    'Ignored rows marked: ' + ignored + '\n' +
    'Skipped rows: ' + skipped
  );
}

/////////////////////////////////////
// CLEAR PDF REVIEW SHEET
/////////////////////////////////////

function clearPdfReviewSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PDF_REVIEW_SHEET_NAME_);

  if (!sheet) throw new Error('Sheet "PDF Review" not found.');

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Clear PDF Review?',
    'This will clear review rows but keep the headers.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  if (sheet.getLastRow() > 1) {
    sheet
      .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
      .clearContent();
  }

  ui.alert('PDF Review cleared.');
}

/////////////////////////////////////
// APPLY PDF REVIEW CONDITIONAL FORMATTING
/////////////////////////////////////
/////////////////////////////////////
// APPLY PDF REVIEW CONDITIONAL FORMATTING
/////////////////////////////////////

function applyPdfReviewConditionalFormatting_(sheet) {
  const headers = getHeaderMap_(sheet, 1);

  const dataLastRow = Math.max(sheet.getLastRow(), 2);
  const formatRows = Math.max(dataLastRow - 1, 1);

  const correctedCasesCol = getRequiredHeader_(headers, 'Corrected Cases', 'PDF Review');
  const correctedUnitsCol = getRequiredHeader_(headers, 'Corrected Units / Weight', 'PDF Review');

  const requiredCols = [
  'Corrected Description',
  'Corrected Pack Size',
  'Corrected Unit Price',
  'Corrected Line Total'
  ];

  const rules = [];

  /////////////////////////////////////
  // CASES / UNITS SPECIAL RULE
  /////////////////////////////////////

  const casesLetter = pdfReviewColumnToLetter_(correctedCasesCol);
  const unitsLetter = pdfReviewColumnToLetter_(correctedUnitsCol);

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=AND(ISBLANK($' + casesLetter + '2),ISBLANK($' + unitsLetter + '2))'
      )
      .setBackground('#f4cccc')
      .setRanges([
        sheet.getRange(2, correctedCasesCol, formatRows, 1),
        sheet.getRange(2, correctedUnitsCol, formatRows, 1)
      ])
      .build()
  );

  /////////////////////////////////////
  // NORMAL REQUIRED FIELDS
  /////////////////////////////////////

  requiredCols.forEach(function(headerName) {
    const col = getRequiredHeader_(headers, headerName, 'PDF Review');
    const letter = pdfReviewColumnToLetter_(col);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISBLANK($' + letter + '2)')
        .setBackground('#f4cccc')
        .setRanges([
          sheet.getRange(2, col, formatRows, 1)
        ])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
}

/////////////////////////////////////
// PDF REVIEW COLUMN LETTER HELPER
/////////////////////////////////////

function pdfReviewColumnToLetter_(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}


/////////////////////////////////////
// CLEAN PDF REVIEW NOTES
/////////////////////////////////////

function cleanPdfReviewNotes_(notes) {
  return (notes || '')
    .toString()
    .replace(/Missing:\s*Item Code/ig, '')
    .replace(/\s*\|\s*\|+\s*/g, ' | ')
    .replace(/^\s*\|\s*/g, '')
    .replace(/\s*\|\s*$/g, '')
    .trim();
}