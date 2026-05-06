/////////////////////////////////////
// BUILD EXTRACTED LINES FROM PDF JSON
/////////////////////////////////////

function buildExtractedLinesFromPdfJson_(fileId) {
  const ss = SpreadsheetApp.getActive();

  if (!fileId) throw new Error('Missing fileId.');

  const json = rebuildJsonFromChunks_(fileId);
  const meta = getPdfJsonMetaByFileId_(fileId);

  const invoiceHeader =
  json.invoiceHeader ||
  json.invoice_header ||
  {};

const resolvedSite =
  json.site ||
  invoiceHeader.siteName ||
  invoiceHeader.site_name ||
  meta.site ||
  '';

  const supplier = (json.supplier || meta.supplier || '').toString().trim().toLowerCase();

  let rows = [];
  let sourceType = '';

  if (supplier === 'bidfood') {
    rows = json.bidfoodRows || [];
    sourceType = 'bidfoodRows';
  } else if (supplier === 'pilgrim') {
    rows = json.pilgrimRows || [];
    sourceType = 'pilgrimRows';
  } else {
    SpreadsheetApp.getUi().alert('Unsupported supplier in JSON: ' + supplier);
    return 0;
  }

  if (!rows.length) {
    SpreadsheetApp.getUi().alert(
      'No invoice rows found in JSON for supplier: ' + (json.supplier || meta.supplier || '')
    );
    return 0;
  }

  const sheet = getOrCreatePdfExtractedLinesSheet_();
  clearExtractedLinesForFile_(sheet, fileId);

const output = rows.map((row, index) => {
  return [
    meta && meta.uploadTime ? meta.uploadTime : new Date(),
    json.fileName || (meta && meta.fileName) || '',
    json.supplier || (meta && meta.supplier) || '',
    resolvedSite,
    fileId,
    index + 1,
    index + 1,
    sourceType,
    '',
    '',
    row.cases || '',
    row.units_weight || '',
    row.base_unit || '',
    row.description || '',
    row.pack_size || '',
    row.item_code || '',
    row.unit_price || '',
    row.line_total || '',
    row.vat || row.vat_rate || '',
    row.vat_total || '',
    row.reviewFlag || ''
  ];
});

  const startRow = Math.max(sheet.getLastRow() + 1, 2);

  sheet
    .getRange(startRow, 1, output.length, output[0].length)
    .setValues(output);

  return output.length;
}


/////////////////////////////////////
// RUN HEADER + EXTRACTED LINES + REVIEW
// ONE MENU ACTION
/////////////////////////////////////

function runBuildExtractedLinesFromPdfJson() {
  const ui = SpreadsheetApp.getUi();

  try {
    const fileId = getLatestPdfJsonDriveFileIdForPipeline_();

    if (!fileId) {
      ui.alert('No latest PDF JSON file found in PDF JSON Staging.');
      return;
    }

    /////////////////////////////////////
    // 1. BUILD / UPDATE PDF INVOICE HEADER
    /////////////////////////////////////

    buildPdfInvoiceHeaderFromLatestJson_(fileId);

    /////////////////////////////////////
    // 2. BUILD PDF EXTRACTED LINES
    /////////////////////////////////////

    buildExtractedLinesFromPdfJson_(fileId);

    /////////////////////////////////////
    // 3. BUILD PDF REVIEW SILENTLY
    /////////////////////////////////////

    buildPdfReviewFromExtractedLines(fileId, { silent: true });

    /////////////////////////////////////
    // 4. FINAL POPUP BASED ON ACTUAL REVIEW SHEET
    /////////////////////////////////////

    const reviewResult = getPdfReviewPopupResultForFile_(fileId);

    const reviewCount = reviewResult.reviewCount;
    const popupLines = reviewResult.popupLines;

    if (reviewCount > 0) {
      ui.alert(
        'PDF build complete.\n\n' +
        'Done:\n' +
        '1. PDF Invoice Headers updated\n' +
        '2. PDF Extracted Lines built\n' +
        '3. PDF Review built\n\n' +
        'Rows needing review: ' + reviewCount + '\n\n' +
        popupLines.slice(0, 15).join('\n\n') +
        (popupLines.length > 15 ? '\n\nMore rows exist in PDF Review.' : '')
      );
    } else {
      ui.alert(
        'PDF build complete.\n\n' +
        'Done:\n' +
        '1. PDF Invoice Headers updated\n' +
        '2. PDF Extracted Lines built\n' +
        '3. PDF Review built\n\n' +
        'No rows need review.'
      );
    }

  } catch (err) {
    ui.alert(
      'PDF build failed:\n\n' +
      (err && err.message ? err.message : err)
    );
    throw err;
  }
}


/////////////////////////////////////
// GET PDF REVIEW POPUP RESULT FOR FILE
// COUNTS ACTUAL PENDING REVIEW ROWS
/////////////////////////////////////

function getPdfReviewPopupResultForFile_(fileId) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Review');

  if (!sheet) {
    return {
      reviewCount: 0,
      popupLines: []
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      reviewCount: 0,
      popupLines: []
    };
  }

  const headers = getHeaderMap_(sheet, 1);

  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Review');
  const statusCol = getRequiredHeader_(headers, 'Review Status', 'PDF Review');

  const rowNoCol = getOptionalHeader_(headers, 'Row No');
  const descCol = getOptionalHeader_(headers, 'Corrected Description');
  const packSizeCol = getOptionalHeader_(headers, 'Corrected Pack Size');
  const itemCodeCol = getOptionalHeader_(headers, 'Corrected Item Code');
  const notesCol = getOptionalHeader_(headers, 'Notes');

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  const popupLines = [];

  data.forEach(row => {
    const rowFileId = row[fileIdCol - 1]
      ? row[fileIdCol - 1].toString().trim()
      : '';

    if (rowFileId !== fileId.toString().trim()) return;

    const status = row[statusCol - 1]
      ? row[statusCol - 1].toString().trim()
      : '';

    if (status !== 'Pending' && status !== 'Needs Cloud Fix') return;

    const rowNo = rowNoCol ? row[rowNoCol - 1] : '';
    const desc = descCol ? row[descCol - 1] : '';
    const packSize = packSizeCol ? row[packSizeCol - 1] : '';
    const itemCode = itemCodeCol ? row[itemCodeCol - 1] : '';
    const notes = notesCol ? row[notesCol - 1] : '';

    popupLines.push(
      'Row ' + rowNo + ': ' +
      desc +
      (packSize ? '\nPack Size: ' + packSize : '') +
      (itemCode ? '\nItem Code: ' + itemCode : '') +
      (notes ? '\nNotes: ' + notes : '')
    );
  });

  return {
    reviewCount: popupLines.length,
    popupLines: popupLines
  };
}


/////////////////////////////////////
// GET LATEST PDF JSON DRIVE FILE ID
// HEADER-BASED
/////////////////////////////////////

function getLatestPdfJsonDriveFileIdForPipeline_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF JSON Staging');

  if (!sheet) {
    throw new Error('Sheet "PDF JSON Staging" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const headers = getHeaderMap_(sheet, 1);
  const driveFileIdCol = getRequiredHeader_(headers, 'Drive File ID');

  const values = sheet
    .getRange(2, driveFileIdCol, lastRow - 1, 1)
    .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const fileId = values[i][0];

    if (fileId) {
      return fileId.toString().trim();
    }
  }

  return '';
}


/////////////////////////////////////
// PDF EXTRACTED LINES HEADERS
/////////////////////////////////////

function getPdfExtractedLinesHeaders_() {
  return [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'Row No',
    'Line No',
    'Source Type',
    'Source Start Line',
    'Source End Line',
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
  ];
}


/////////////////////////////////////
// GET OR CREATE PDF EXTRACTED LINES SHEET
/////////////////////////////////////

function getOrCreatePdfExtractedLinesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('PDF Extracted Lines');

  if (!sheet) {
    sheet = ss.insertSheet('PDF Extracted Lines');
  }

  const headers = getPdfExtractedLinesHeaders_();

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  return sheet;
}


/////////////////////////////////////
// CLEAR EXTRACTED LINES FOR FILE
/////////////////////////////////////

function clearExtractedLinesForFile_(sheet, fileId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headerMap = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headerMap, 'Drive File ID', 'PDF Extracted Lines');

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
// TEST BUILD EXTRACTED LINES ONLY
/////////////////////////////////////

function testBuildExtractedLinesFromPdfJson() {
  const fileId = Browser.inputBox('Enter Drive File ID to build Extracted Lines');

  if (!fileId || fileId === 'cancel') return;

  const count = buildExtractedLinesFromPdfJson_(fileId);

  SpreadsheetApp.getUi().alert(
    'PDF Extracted Lines built successfully.\n\n' +
    'Rows written: ' + count
  );
}


/////////////////////////////////////
// TEST BUILD EXTRACTED LINES + REVIEW
/////////////////////////////////////

function testBuildExtractedLinesAndReviewFromPdfJson() {
  const fileId = Browser.inputBox('Enter Drive File ID to build Extracted Lines + Review');

  if (!fileId || fileId === 'cancel') return;

  const extractedCount = buildExtractedLinesFromPdfJson_(fileId);

  let reviewCount = 0;

  if (typeof buildPdfReviewFromExtractedLines === 'function') {
    buildPdfReviewFromExtractedLines(fileId, { silent: true });

    const reviewResult = getPdfReviewPopupResultForFile_(fileId);

    reviewCount = reviewResult.reviewCount;
  }

  SpreadsheetApp.getUi().alert(
    'PDF Extracted Lines + Review built successfully.\n\n' +
    'Extracted rows: ' + extractedCount + '\n' +
    'Rows needing review: ' + reviewCount
  );
}


/////////////////////////////////////
// HEADER VALUE HELPERS
/////////////////////////////////////

function getValueByHeader_(row, headerMap, headerName) {
  const col = getRequiredHeader_(headerMap, headerName, 'Header lookup');
  return row[col - 1];
}


function getOptionalValueByHeader_(row, headerMap, headerName) {
  const col = headerMap[headerName];
  return col ? row[col - 1] : '';
}