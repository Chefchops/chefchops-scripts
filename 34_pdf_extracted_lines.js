/////////////////////////////////////
// BUILD EXTRACTED LINES FROM PDF JSON
/////////////////////////////////////

function buildExtractedLinesFromPdfJson_(fileId) {
  const ss = SpreadsheetApp.getActive();

  if (!fileId) throw new Error('Missing fileId.');

  const json = rebuildJsonFromChunks_(fileId);
  const meta = getPdfJsonMetaByFileId_(fileId);

  const rows = json.bidfoodRows || [];

  if (!rows.length) {
    SpreadsheetApp.getUi().alert('No bidfoodRows found in JSON.');
    return 0;
  }

  const sheet = getOrCreatePdfExtractedLinesSheet_();
  clearExtractedLinesForFile_(sheet, fileId);

  const output = rows.map((row, index) => {
    return [
      meta && meta.uploadTime ? meta.uploadTime : new Date(),
      json.fileName || (meta && meta.fileName) || '',
      json.supplier || (meta && meta.supplier) || '',
      json.site || (meta && meta.site) || '',
      fileId,
      index + 1,
      index + 1,
      'bidfoodRows',
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
      row.vat || '',
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
// RUN BUILD EXTRACTED LINES + REVIEW
/////////////////////////////////////

function runBuildExtractedLinesFromPdfJson() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const stagingSheet = ss.getSheetByName('PDF Staging');

  if (!stagingSheet) {
    ui.alert('Missing sheet: PDF Staging');
    return;
  }

  const headers = getHeaderMap_(stagingSheet, 1);

  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');
  const apiStatusCol = getRequiredHeader_(headers, 'API Status', 'PDF Staging');

  const lastRow = stagingSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF Staging rows found.');
    return;
  }

  const fileId = stagingSheet.getRange(lastRow, fileIdCol).getValue();
  const apiStatus = stagingSheet.getRange(lastRow, apiStatusCol).getValue();

  if (!fileId) {
    ui.alert('Missing Drive File ID on latest PDF Staging row.');
    return;
  }

  if (apiStatus !== 'DONE') {
    ui.alert('Latest PDF has not completed Cloud processing yet.');
    return;
  }

  const extractedCount = buildExtractedLinesFromPdfJson_(fileId);

  let reviewCount = 0;

  if (extractedCount > 0) {
    if (typeof buildPdfReviewFromExtractedLines !== 'function') {
      ui.alert(
        'PDF Extracted Lines built, but PDF Review was not built.\n\n' +
        'Missing function:\n' +
        'buildPdfReviewFromExtractedLines(fileId)'
      );
      return;
    }

    reviewCount = buildPdfReviewFromExtractedLines(fileId) || 0;
  }

   if (reviewCount === 0) {
    ui.alert(
      'PDF Extracted Lines complete.\n\n' +
      'Drive File ID:\n' + fileId + '\n\n' +
      'Extracted rows: ' + extractedCount + '\n' +
      'No review rows needed.'
    );
  }
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
    reviewCount = buildPdfReviewFromExtractedLines(fileId) || 0;
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