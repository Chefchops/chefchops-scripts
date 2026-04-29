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
    base64Pdf: base64
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
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
      raw: resultText
    };
  }

  stagingSheet.getRange(rowNumber, 6).setValue(statusCode === 200 ? 'DONE' : 'ERROR'); // API Status
  stagingSheet.getRange(rowNumber, 7).setValue(parsed.success ? 'STORED' : 'FAILED');   // JSON Status
  stagingSheet.getRange(rowNumber, 8).setValue(
    parsed.success
      ? 'JSON stored in PDF JSON Staging'
      : (parsed.error || 'Unknown error')
  ); // Notes

  clearJsonChunksForFile_(jsonSheet, fileId);

  writeJsonToStaging_(jsonSheet, {
    uploadTime: uploadTime,
    fileName: fileName,
    supplier: supplier,
    site: site,
    fileId: fileId,
    jsonText: resultText
  });
}

/////////////////////////////////////
// BUILD EXTRACTED LINES FOR LAST PDF ROW
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

  const fileId = stagingSheet.getRange(lastRow, 5).getValue(); // E = Drive File ID
  if (!fileId) {
    ui.alert('No Drive File ID found on the last PDF row.');
    return;
  }

  const result = buildExtractedLinesFromPdfJson_(fileId);

  ui.alert(
    'Built extracted lines for file ID: ' + fileId +
    '\n\nLines written: ' + (result && result.extractedLines ? result.extractedLines : 'Done')
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

  const fileId = stagingSheet.getRange(lastRow, 5).getValue(); // E = Drive File ID
  if (!fileId) {
    ui.alert('No Drive File ID found on the last PDF row.');
    return;
  }

  const result = buildParsedRowsFromExtractedLines_(fileId);

  ui.alert(
    'Built parsed rows for file ID: ' + fileId +
    '\n\nRows written: ' + (result && result.rowsWritten ? result.rowsWritten : 'Done')
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
    ui.alert('No PDF parsed rows found for ' + supplier + (site ? ' / ' + site : '') + '.');
    return;
  }

  clearInvoiceImportSilent_();

  invoiceSheet.getRange('B4').setValue(supplier);
  invoiceSheet.getRange('B5').setValue(site);

  const startRow = 8;
  const output = rows.map(function(r) {
    return [
      r.description || '', // A
      r.qty || '',         // B
      r.unit || '',        // C
      r.unitPrice || ''    // D
    ];
  });

  invoiceSheet.getRange(startRow, 1, output.length, 4).setValues(output);

  ui.alert(
    'PDF parsed rows loaded into Invoice Import raw area.\n\n' +
    'Supplier: ' + supplier + '\n' +
    'Rows loaded: ' + output.length + '\n\n' +
    'Now run Build Invoice Import.'
  );
}