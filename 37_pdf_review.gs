/////////////////////////////////////
// PDF REVIEW SYSTEM
// ACTIVE CURRENT PIPELINE
//
// Current flow:
// PDF Extracted Lines
// -> PDF Review
// -> Apply Review Corrections
// -> PDF Extracted Lines
//
// Legacy removed from this file:
// PDF Parsed Rows
// buildPdfReviewSheet()
// parsed-row review gate
/////////////////////////////////////

const PDF_REVIEW_SHEET_NAME_ = 'PDF Review';

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
  const ui = SpreadsheetApp.getUi();

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

  const reviewHeaders = getHeaderMap_(sheet, 1);

  const statusCol = getRequiredHeader_(reviewHeaders, 'Review Status', 'PDF Review');
  const originalItemCodeCol = getRequiredHeader_(reviewHeaders, 'Original Item Code', 'PDF Review');
  const correctedItemCodeCol = getRequiredHeader_(reviewHeaders, 'Corrected Item Code', 'PDF Review');

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Ignore Row', 'Needs Cloud Fix'], true)
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(2, statusCol, Math.max(sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(rule);

  // Keep item codes as text so trailing zeroes are preserved.
  sheet.getRange(2, originalItemCodeCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  sheet.getRange(2, correctedItemCodeCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');

  applyPdfReviewConditionalFormatting_(sheet);

  sheet.autoResizeColumns(1, headers.length);

  ui.alert('PDF Review sheet setup complete.');
}

/////////////////////////////////////
// BUILD PDF REVIEW FROM EXTRACTED LINES
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
  ].forEach(function(headerName) {
    getRequiredHeader_(sourceHeaders, headerName, 'PDF Extracted Lines');
  });

  getPdfReviewHeaders_().forEach(function(headerName) {
    getRequiredHeader_(reviewHeaders, headerName, 'PDF Review');
  });

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
    const sourceFileId = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Drive File ID');

    if ((sourceFileId || '').toString().trim() !== fileId.toString().trim()) return;

    const rowNo = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Row No');
    const reviewFlag = (pdfReviewGetValueByHeader_(row, sourceHeaders, 'Review Flag') || '').toString().trim();

    const cases = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Cases');
    const unitsWeight = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Units / Weight');
    const description = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Description');
    const packSize = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Pack Size');
    const itemCode = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Item Code');
    const unitPrice = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Unit Price');
    const lineTotal = pdfReviewGetValueByHeader_(row, sourceHeaders, 'Line Total');

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
      'File Name': pdfReviewGetValueByHeader_(row, sourceHeaders, 'File Name'),
      'Supplier': pdfReviewGetValueByHeader_(row, sourceHeaders, 'Supplier'),
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
      'Notes': cleanPdfReviewNotes_(notes.join(' | '))
    });

    output.push(reviewRow);

    popupLines.push(
      'Row ' + rowNo + ': ' + (description || '[No description]') + '\n' +
      cleanPdfReviewNotes_(notes.join(' | '))
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

  const reviewStatusCol = getRequiredHeader_(reviewHeaders, 'Review Status', 'PDF Review');
  const correctedItemCodeCol = getRequiredHeader_(reviewHeaders, 'Corrected Item Code', 'PDF Review');
  const originalItemCodeCol = getRequiredHeader_(reviewHeaders, 'Original Item Code', 'PDF Review');

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Ignore Row', 'Needs Cloud Fix'], true)
    .setAllowInvalid(false)
    .build();

  reviewSheet
    .getRange(startRow, reviewStatusCol, output.length, 1)
    .setDataValidation(rule);

  // Preserve item codes as text.
  reviewSheet.getRange(startRow, originalItemCodeCol, output.length, 1).setNumberFormat('@');
  reviewSheet.getRange(startRow, correctedItemCodeCol, output.length, 1).setNumberFormat('@');

  applyPdfReviewConditionalFormatting_(reviewSheet);

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

  fileIds.forEach(function(value, index) {
    if ((value || '').toString().trim() === fileId.toString().trim()) {
      rowsToDelete.push(index + 2);
    }
  });

  if (!rowsToDelete.length) return;

  pdfReviewDeleteRowsInGroups_(sheet, rowsToDelete);
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
// APPLY PDF REVIEW CORRECTIONS
// PDF Review -> PDF Extracted Lines
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

  extractedValues.forEach(function(row, index) {
    const fileId = pdfReviewGetValueByHeader_(row, extractedHeaders, 'Drive File ID');
    const rowNo = pdfReviewGetValueByHeader_(row, extractedHeaders, 'Row No');

    if (fileId && rowNo) {
      extractedLookup[fileId + '|' + rowNo] = index + 2;
    }
  });

  let applied = 0;
  let ignored = 0;
  let skipped = 0;

  reviewValues.forEach(function(reviewRow) {
    const status = (pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Review Status') || '').toString().trim();

    if (status !== 'Approved' && status !== 'Ignore Row') {
      skipped++;
      return;
    }

    const fileId = pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Drive File ID');
    const rowNo = pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Row No');
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
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Cases'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Units / Weight', 'PDF Extracted Lines'))
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Units / Weight'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Description', 'PDF Extracted Lines'))
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Description'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Pack Size', 'PDF Extracted Lines'))
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Pack Size'));

    const itemCodeValue = (pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Item Code') || '').toString();

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Item Code', 'PDF Extracted Lines'))
      .setNumberFormat('@')
      .setValue(itemCodeValue);

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Unit Price', 'PDF Extracted Lines'))
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Unit Price'));

    extractedSheet
      .getRange(targetRow, getRequiredHeader_(extractedHeaders, 'Line Total', 'PDF Extracted Lines'))
      .setValue(pdfReviewGetValueByHeader_(reviewRow, reviewHeaders, 'Corrected Line Total'));

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
  const ui = SpreadsheetApp.getUi();

  if (!sheet) throw new Error('Sheet "PDF Review" not found.');

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

  applyPdfReviewConditionalFormatting_(sheet);

  ui.alert('PDF Review cleared.');
}

/////////////////////////////////////
// HIGHLIGHT PDF REVIEW MISSING FIELDS
/////////////////////////////////////

function highlightPdfReviewMissingFields() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('PDF Review');

  if (!sheet) {
    ui.alert('Sheet "PDF Review" not found.');
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF Review rows to highlight.');
    return;
  }

  applyPdfReviewConditionalFormatting_(sheet);

  ui.alert('PDF Review missing-field highlighting applied.');
}

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

  /////////////////////////////////////
  // REVIEW STATUS COLOURS
  /////////////////////////////////////

  const statusCol = getRequiredHeader_(headers, 'Review Status', 'PDF Review');
  const statusLetter = pdfReviewColumnToLetter_(statusCol);
  const fullRowRange = sheet.getRange(2, 1, formatRows, sheet.getLastColumn());

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusLetter + '2="Pending"')
      .setBackground('#fff2cc')
      .setRanges([fullRowRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusLetter + '2="Approved"')
      .setBackground('#d9ead3')
      .setRanges([fullRowRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusLetter + '2="Needs Cloud Fix"')
      .setBackground('#f4cccc')
      .setRanges([fullRowRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + statusLetter + '2="Ignore Row"')
      .setBackground('#d9d9d9')
      .setRanges([fullRowRange])
      .build()
  );

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
    .replace(/\s+/g, ' ')
    .trim();
}

/////////////////////////////////////
// GET VALUE BY HEADER
// LOCAL TO PDF REVIEW FILE
/////////////////////////////////////

function pdfReviewGetValueByHeader_(row, headerMap, headerName) {
  const col = headerMap[headerName];
  if (!col) return '';
  return row[col - 1];
}

/////////////////////////////////////
// DELETE ROWS IN GROUPS
// LOCAL SAFE HELPER
/////////////////////////////////////

function pdfReviewDeleteRowsInGroups_(sheet, rows) {
  if (!rows || !rows.length) return;

  const sorted = rows
    .map(Number)
    .filter(Boolean)
    .sort(function(a, b) {
      return b - a;
    });

  let groupStart = sorted[0];
  let groupEnd = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const row = sorted[i];

    if (row === groupEnd - 1) {
      groupEnd = row;
    } else {
      sheet.deleteRows(groupEnd, groupStart - groupEnd + 1);
      groupStart = row;
      groupEnd = row;
    }
  }

  sheet.deleteRows(groupEnd, groupStart - groupEnd + 1);
}