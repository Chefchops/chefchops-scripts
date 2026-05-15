/////////////////////////////////////
// PROCESS LAST PDF ROW
/////////////////////////////////////

function processLastPdfRow_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF rows found in PDF Staging.');
    return;
  }

  processPdfRow(lastRow);

  ui.alert('Processed PDF Staging row ' + lastRow + '.');
}

/////////////////////////////////////
// PROCESS NEXT PENDING PDF ROW
/////////////////////////////////////

function processNextPendingPdfRow() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const target = getNextPendingPdfStagingRow_(sheet);

  if (!target.rowNumber) {
    ui.alert('No pending PDF rows found.');
    return;
  }

  processPdfRow(target.rowNumber);

  ui.alert(
    'Processed next pending PDF.\n\n' +
      'Row: ' +
      target.rowNumber +
      '\n' +
      'File: ' +
      target.fileName,
  );
}

/////////////////////////////////////
// PROCESS NEXT 5 PENDING PDF ROWS
/////////////////////////////////////

function processNext5PendingPdfRows() {
  processPendingPdfRows_(5);
}

/////////////////////////////////////
// PROCESS PENDING PDF ROWS
/////////////////////////////////////

function processPendingPdfRows_(limit) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF rows found in PDF Staging.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);

  const statusCol = getRequiredHeader_(headers, 'API Status', 'PDF Staging');
  const jsonStatusCol = getRequiredHeader_(
    headers,
    'JSON Status',
    'PDF Staging',
  );
  const fileNameCol = getRequiredHeader_(headers, 'File Name', 'PDF Staging');
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');

  const values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  let processed = 0;
  let failed = 0;
  const messages = [];

  for (let i = 0; i < values.length; i++) {
    if (processed >= limit) break;

    const row = values[i];
    const sheetRow = i + 2;

    const apiStatus = normalisePdfStatus_(row[statusCol - 1]);
    const jsonStatus = normalisePdfStatus_(row[jsonStatusCol - 1]);
    const fileName = row[fileNameCol - 1] || '';
    const fileId = (row[fileIdCol - 1] || '').toString().trim();

    if (!fileId) continue;

    if (apiStatus === 'DONE' && jsonStatus !== 'FAILED') continue;

    if (
      apiStatus &&
      apiStatus !== 'PENDING' &&
      apiStatus !== 'FAILED' &&
      apiStatus !== 'ERROR' &&
      apiStatus !== 'RETRY'
    ) {
      continue;
    }

    try {
      sheet.getRange(sheetRow, statusCol).setValue('PROCESSING');

      processPdfRow(sheetRow);

      processed++;
      messages.push('OK row ' + sheetRow + ': ' + fileName);
    } catch (err) {
      failed++;

      sheet.getRange(sheetRow, statusCol).setValue('ERROR');
      sheet.getRange(sheetRow, jsonStatusCol).setValue('FAILED');

      messages.push(
        'FAILED row ' + sheetRow + ': ' + fileName + ' | ' + err.message,
      );
    }
  }

  ui.alert(
    'PDF batch processing complete.\n\n' +
      'Limit: ' +
      limit +
      '\n' +
      'Processed: ' +
      processed +
      '\n' +
      'Failed: ' +
      failed +
      '\n\n' +
      messages.slice(0, 15).join('\n') +
      (messages.length > 15 ? '\n\nMore results not shown.' : ''),
  );
}

/////////////////////////////////////
// GET NEXT PENDING PDF STAGING ROW
/////////////////////////////////////

function getNextPendingPdfStagingRow_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      rowNumber: 0,
      fileName: '',
    };
  }

  const headers = getHeaderMap_(sheet, 1);

  const statusCol = getRequiredHeader_(headers, 'API Status', 'PDF Staging');
  const fileNameCol = getRequiredHeader_(headers, 'File Name', 'PDF Staging');
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');

  const values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const sheetRow = i + 2;

    const apiStatus = normalisePdfStatus_(row[statusCol - 1]);
    const fileId = (row[fileIdCol - 1] || '').toString().trim();

    if (!fileId) continue;

    if (
      apiStatus === '' ||
      apiStatus === 'PENDING' ||
      apiStatus === 'FAILED' ||
      apiStatus === 'ERROR' ||
      apiStatus === 'RETRY'
    ) {
      return {
        rowNumber: sheetRow,
        fileName: row[fileNameCol - 1] || '',
      };
    }
  }

  return {
    rowNumber: 0,
    fileName: '',
  };
}

/////////////////////////////////////
// PROCESS ONE PDF ROW
/////////////////////////////////////

function processPdfRow(rowNumber) {
  const ss = SpreadsheetApp.getActive();
  const stagingSheet = ss.getSheetByName('PDF Staging');
  const jsonSheet = ss.getSheetByName('PDF JSON Staging');

  if (!stagingSheet) throw new Error('Sheet "PDF Staging" not found.');
  if (!jsonSheet) throw new Error('Sheet "PDF JSON Staging" not found.');

  if (!rowNumber || rowNumber < 2) {
    throw new Error('Please pass a valid row number, e.g. processPdfRow(2).');
  }

  const data = stagingSheet.getRange(rowNumber, 1, 1, 8).getValues()[0];

  const uploadTime = data[0];
  const fileName = data[1];
  const supplier = data[2];
  const site = data[3];
  const fileId = data[4];

  if (!fileId) {
    throw new Error('No Drive File ID found on row ' + rowNumber);
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  const url = 'https://chefchops-pdf-parser-639314070996.europe-west1.run.app';

  const payload = {
    fileName: fileName,
    supplier: supplier,
    site: site,
    base64Pdf: base64,
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const resultText = response.getContentText();
  const statusCode = response.getResponseCode();

  let parsed;

  try {
    parsed = JSON.parse(resultText);
  } catch (err) {
    parsed = {
      success: false,
      error: 'Response was not valid JSON',
      raw: resultText,
    };
  }

  stagingSheet
    .getRange(rowNumber, 6)
    .setValue(statusCode === 200 ? 'DONE' : 'ERROR');

  stagingSheet
    .getRange(rowNumber, 7)
    .setValue(parsed.success ? 'STORED' : 'FAILED');

  stagingSheet
    .getRange(rowNumber, 8)
    .setValue(
      parsed.success
        ? 'JSON stored in PDF JSON Staging'
        : parsed.error || 'Unknown error',
    );

  clearJsonChunksForFile_(jsonSheet, fileId);

  writeJsonToStaging_(jsonSheet, {
    uploadTime: uploadTime,
    fileName: fileName,
    supplier: supplier,
    site: site,
    fileId: fileId,
    jsonText: resultText,
  });
}

/////////////////////////////////////
// BUILD NEXT 5 STORED PDF JSON FILES
/////////////////////////////////////

function buildNext5StoredPdfJsonFiles() {
  buildStoredPdfJsonFiles_(5);
}

/////////////////////////////////////
// BUILD STORED PDF JSON FILES
/////////////////////////////////////

function buildStoredPdfJsonFiles_(limit) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const stagingSheet = ss.getSheetByName('PDF Staging');

  if (!stagingSheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const lastRow = stagingSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF rows found in PDF Staging.');
    return;
  }

  const headers = getHeaderMap_(stagingSheet, 1);

  const apiStatusCol = getRequiredHeader_(headers, 'API Status', 'PDF Staging');
  const jsonStatusCol = getRequiredHeader_(
    headers,
    'JSON Status',
    'PDF Staging',
  );
  const fileNameCol = getRequiredHeader_(headers, 'File Name', 'PDF Staging');
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');
  const notesCol = getRequiredHeader_(headers, 'Notes', 'PDF Staging');

  const values = stagingSheet
    .getRange(2, 1, lastRow - 1, stagingSheet.getLastColumn())
    .getValues();

  let built = 0;
  let failed = 0;
  const messages = [];

  for (let i = 0; i < values.length; i++) {
    if (built >= limit) break;

    const row = values[i];
    const sheetRow = i + 2;

    const apiStatus = normalisePdfStatus_(row[apiStatusCol - 1]);
    const jsonStatus = normalisePdfStatus_(row[jsonStatusCol - 1]);
    const fileName = row[fileNameCol - 1] || '';
    const fileId = (row[fileIdCol - 1] || '').toString().trim();

    if (!fileId) continue;
    if (apiStatus !== 'DONE') continue;
    if (jsonStatus !== 'STORED') continue;

    try {
      stagingSheet.getRange(sheetRow, jsonStatusCol).setValue('BUILDING');

      const result = buildPdfHeaderExtractedLinesAndReviewForFile_(fileId);

      if (result.reviewCount > 0) {
        stagingSheet
          .getRange(sheetRow, jsonStatusCol)
          .setValue('REVIEW NEEDED');
        stagingSheet
          .getRange(sheetRow, notesCol)
          .setValue(
            'Header + Extracted Lines + Review built | Rows needing review: ' +
              result.reviewCount,
          );
      } else {
        stagingSheet
          .getRange(sheetRow, jsonStatusCol)
          .setValue('READY TO APPEND');
        stagingSheet
          .getRange(sheetRow, notesCol)
          .setValue('Header + Extracted Lines + Review built | No review rows');
      }

      built++;

      messages.push(
        'OK row ' +
          sheetRow +
          ': ' +
          fileName +
          ' | Lines: ' +
          (result && result.extractedCount ? result.extractedCount : 'Done') +
          ' | Review: ' +
          (result && result.reviewCount ? result.reviewCount : 0),
      );
    } catch (err) {
      failed++;

      stagingSheet.getRange(sheetRow, jsonStatusCol).setValue('FAILED');
      stagingSheet
        .getRange(sheetRow, notesCol)
        .setValue('Build failed: ' + err.message);

      messages.push(
        'FAILED row ' + sheetRow + ': ' + fileName + ' | ' + err.message,
      );
    }
  }

  ui.alert(
    'PDF JSON build batch complete.\n\n' +
      'Limit: ' +
      limit +
      '\n' +
      'Built: ' +
      built +
      '\n' +
      'Failed: ' +
      failed +
      '\n\n' +
      messages.slice(0, 15).join('\n') +
      (messages.length > 15 ? '\n\nMore results not shown.' : ''),
  );
}

/////////////////////////////////////
// BUILD HEADER + EXTRACTED LINES + REVIEW FOR LAST PDF ROW
/////////////////////////////////////

function buildExtractedLinesForLastPdfRow_() {
  const ss = SpreadsheetApp.getActive();
  const stagingSheet = ss.getSheetByName('PDF Staging');
  const ui = SpreadsheetApp.getUi();

  if (!stagingSheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const lastRow = stagingSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF rows found in PDF Staging.');
    return;
  }

  const fileId = stagingSheet.getRange(lastRow, 5).getValue();

  if (!fileId) {
    ui.alert('No Drive File ID found on the last PDF row.');
    return;
  }

  const result = buildPdfHeaderExtractedLinesAndReviewForFile_(fileId);

  ui.alert(
    'Built Header + Extracted Lines + Review for file ID: ' +
      fileId +
      '\n\nExtracted rows: ' +
      (result && result.extractedCount ? result.extractedCount : 'Done') +
      '\nRows needing review: ' +
      (result && result.reviewCount ? result.reviewCount : 0),
  );
}

/////////////////////////////////////
// BUILD PARSED ROWS FOR LAST PDF ROW
/////////////////////////////////////

function buildParsedRowsForLastPdfRow_() {
  const ss = SpreadsheetApp.getActive();
  const stagingSheet = ss.getSheetByName('PDF Staging');
  const ui = SpreadsheetApp.getUi();

  if (!stagingSheet) {
    ui.alert('Sheet "PDF Staging" not found.');
    return;
  }

  const lastRow = stagingSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No PDF rows found in PDF Staging.');
    return;
  }

  const fileId = stagingSheet.getRange(lastRow, 5).getValue();

  if (!fileId) {
    ui.alert('No Drive File ID found on the last PDF row.');
    return;
  }

  const result = buildParsedRowsFromExtractedLines_(fileId);

  ui.alert(
    'Built parsed rows for file ID: ' +
      fileId +
      '\n\nRows written: ' +
      (result && result.rowsWritten ? result.rowsWritten : 'Done'),
  );
}

/////////////////////////////////////
// LOAD PDF PARSED ROWS TO INVOICE IMPORT RAW
/////////////////////////////////////

function loadPdfLinesToInvoiceImportRaw_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const invoiceSheet = ss.getSheetByName('Invoice Import');

  if (!invoiceSheet) {
    ui.alert('Missing "Invoice Import" sheet.');
    return;
  }

  const context = getConfirmedInvoiceContext();

  if (!context) return;

  const supplier = (context.supplier || '').toString().trim();
  const site = (context.site || '').toString().trim();

  const rows = getPdfParsedRowsForContext_(supplier, site);

  if (!rows.length) {
    ui.alert(
      'No PDF parsed rows found for ' +
        supplier +
        (site ? ' / ' + site : '') +
        '.',
    );
    return;
  }

  clearInvoiceImportSilent_();

  invoiceSheet.getRange('B4').setValue(supplier);
  invoiceSheet.getRange('B5').setValue(site);

  const startRow = 8;

  const output = rows.map(function (r) {
    return [r.description || '', r.qty || '', r.unit || '', r.unitPrice || ''];
  });

  invoiceSheet.getRange(startRow, 1, output.length, 4).setValues(output);

  ui.alert(
    'PDF parsed rows loaded into Invoice Import raw area.\n\n' +
      'Supplier: ' +
      supplier +
      '\n' +
      'Rows loaded: ' +
      output.length +
      '\n\n' +
      'Now run Build Invoice Import.',
  );
}

/////////////////////////////////////
// STATUS HELPER
/////////////////////////////////////

function normalisePdfStatus_(value) {
  return (value || '').toString().trim().toUpperCase();
}
