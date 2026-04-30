/////////////////////////////////////
// BUILD INVOICE IMPORT FROM PDF REVIEW
/////////////////////////////////////

function buildInvoiceImportFromPdfReview() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const reviewSheet = ss.getSheetByName('PDF Review');
  const importSheet = ss.getSheetByName('Invoice Import');

  if (!reviewSheet) {
    ui.alert('Missing sheet: PDF Review');
    return;
  }

  if (!importSheet) {
    ui.alert('Missing sheet: Invoice Import');
    return;
  }

  const reviewData = reviewSheet.getDataRange().getValues();

  if (reviewData.length < 2) {
    ui.alert('PDF Review has no rows.');
    return;
  }

  const reviewHeaders = getHeaderMap_(reviewSheet, 1);
  const importHeaders = getHeaderMap_(importSheet, 1);

  const rowsToWrite = [];

  for (let r = 1; r < reviewData.length; r++) {
    const row = reviewData[r];

    const status = getCellByHeader_(row, reviewHeaders, 'Review Status');

    if (status !== 'Approved') continue;

    const supplier = getCellByHeader_(row, reviewHeaders, 'Supplier');
    const fileName = getCellByHeader_(row, reviewHeaders, 'File Name');
    const fileId = getCellByHeader_(row, reviewHeaders, 'Drive File ID');
    const rowNo = getCellByHeader_(row, reviewHeaders, 'Row No');

    const cases = pickCorrectedOrOriginal_(row, reviewHeaders, 'Cases');
    const units = pickCorrectedOrOriginal_(row, reviewHeaders, 'Units / Weight');
    const description = pickCorrectedOrOriginal_(row, reviewHeaders, 'Description');
    const packSize = pickCorrectedOrOriginal_(row, reviewHeaders, 'Pack Size');
    const itemCode = pickCorrectedOrOriginal_(row, reviewHeaders, 'Item Code');
    const unitPrice = pickCorrectedOrOriginal_(row, reviewHeaders, 'Unit Price');
    const lineTotal = pickCorrectedOrOriginal_(row, reviewHeaders, 'Line Total');

    if (!description) continue;

    const outputRow = new Array(importSheet.getLastColumn()).fill('');

    setCellByHeader_(outputRow, importHeaders, 'Supplier', supplier);
    setCellByHeader_(outputRow, importHeaders, 'Cases', cases);
    setCellByHeader_(outputRow, importHeaders, 'Units / Weight', units);
    setCellByHeader_(outputRow, importHeaders, 'Description', description);
    setCellByHeader_(outputRow, importHeaders, 'Pack Size', packSize);
    setCellByHeader_(outputRow, importHeaders, 'Item Code', itemCode);
    setCellByHeader_(outputRow, importHeaders, 'Unit Price', unitPrice);
    setCellByHeader_(outputRow, importHeaders, 'Line Total', lineTotal);
    setCellByHeader_(outputRow, importHeaders, 'Review Flag', 'OK');

    setCellByHeader_(
      outputRow,
      importHeaders,
      'Notes',
      'From PDF Review | File: ' + fileName + ' | Row: ' + rowNo + ' | ID: ' + fileId
    );

    rowsToWrite.push(outputRow);
  }

  if (!rowsToWrite.length) {
    ui.alert('No Approved rows found.');
    return;
  }

  clearInvoiceImportBody_(importSheet);

  importSheet
    .getRange(2, 1, rowsToWrite.length, rowsToWrite[0].length)
    .setValues(rowsToWrite);

  ui.alert('Invoice Import built.\nRows: ' + rowsToWrite.length);
}

/////////////////////////////////////
// HELPERS
/////////////////////////////////////

function pickCorrectedOrOriginal_(row, headers, field) {
  const corrected = getCellByHeader_(row, headers, 'Corrected ' + field);
  const original = getCellByHeader_(row, headers, 'Original ' + field);
  return corrected !== '' ? corrected : original;
}

function getCellByHeader_(row, headers, name) {
  const col = headers[name];
  if (!col) throw new Error('Missing header: ' + name);
  return row[col - 1];
}

function setCellByHeader_(row, headers, name, value) {
  const col = headers[name];
  if (!col) throw new Error('Missing header: ' + name);
  row[col - 1] = value;
}

function clearInvoiceImportBody_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}