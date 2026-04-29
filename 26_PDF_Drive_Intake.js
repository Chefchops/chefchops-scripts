/////////////////////////////////////
// PDF DRIVE INTAKE
// Drive Folder -> PDF Staging
/////////////////////////////////////

const PDF_DRIVE_FOLDER_ID_ = '1uSoRK8QrqjRSXoBBrJWDDn2S576Js-Mv';
const PDF_STAGING_SHEET_NAME_ = 'PDF Staging';

function importPdfJobsFromDriveFolder() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const stagingSheet = ss.getSheetByName(PDF_STAGING_SHEET_NAME_);

  if (!stagingSheet) {
    throw new Error('Missing required sheet: ' + PDF_STAGING_SHEET_NAME_);
  }

  const headers = getHeaderMap_(stagingSheet, 1);

  [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'API Status',
    'JSON Status',
    'Notes'
  ].forEach(h => getRequiredHeader_(headers, h, PDF_STAGING_SHEET_NAME_));

  let folder;

  try {
    folder = DriveApp.getFolderById(PDF_DRIVE_FOLDER_ID_);
  } catch (err) {
    throw new Error('Could not open Drive folder: ' + err.message);
  }

  const files = folder.getFiles();
  const existingIds = getExistingPdfDriveFileIds_(stagingSheet, headers);

  const output = [];
  const now = new Date();

  let scanned = 0;
  let added = 0;
  let skippedNonPdf = 0;
  let skippedExisting = 0;

  while (files.hasNext()) {
    const file = files.next();
    scanned++;

    if (file.getMimeType() !== MimeType.PDF) {
      skippedNonPdf++;
      continue;
    }

    const fileId = file.getId();

    if (existingIds[fileId]) {
      skippedExisting++;
      continue;
    }

    const fileName = file.getName();
    const inferred = inferSupplierAndSiteFromPdfName_(fileName);

    const row = new Array(stagingSheet.getLastColumn()).fill('');

    setRowByHeaders_(row, headers, {
      'Upload Time': now,
      'File Name': fileName,
      'Supplier': inferred.supplier,
      'Site': inferred.site,
      'Drive File ID': fileId,
      'API Status': 'PENDING',
      'JSON Status': '',
      'Notes': inferred.notes
    });

    output.push(row);
    existingIds[fileId] = true;
    added++;
  }

  if (output.length) {
    const startRow = Math.max(stagingSheet.getLastRow() + 1, 2);

    stagingSheet
      .getRange(startRow, 1, output.length, output[0].length)
      .setValues(output);
  }

  const message =
    'PDF Drive intake complete.\n\n' +
    'Files scanned: ' + scanned + '\n' +
    'New PDF jobs added: ' + added + '\n' +
    'Skipped existing: ' + skippedExisting + '\n' +
    'Skipped non-PDF: ' + skippedNonPdf;

  ui.alert(message);
  Logger.log(message);

  return {
    scanned: scanned,
    added: added,
    skippedExisting: skippedExisting,
    skippedNonPdf: skippedNonPdf
  };
}

/////////////////////////////////////
// TEST PDF DRIVE INTAKE
/////////////////////////////////////

function testImportPdfJobsFromDriveFolder() {
  importPdfJobsFromDriveFolder();
}

/////////////////////////////////////
// EXISTING FILE IDS
/////////////////////////////////////

function getExistingPdfDriveFileIds_(sheet, headerMap) {
  const col = getRequiredHeader_(headerMap, 'Drive File ID', PDF_STAGING_SHEET_NAME_);
  const lastRow = sheet.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sheet
    .getRange(2, col, lastRow - 1, 1)
    .getValues();

  values.forEach(row => {
    const id = (row[0] || '').toString().trim();
    if (id) map[id] = true;
  });

  return map;
}

/////////////////////////////////////
// INFER SUPPLIER / SITE FROM FILE NAME
/////////////////////////////////////

function inferSupplierAndSiteFromPdfName_(fileName) {
  const name = (fileName || '').toString().trim().toLowerCase();

  let supplier = '';
  let site = '';
  const notes = [];

  if (name.includes('bidfood')) supplier = 'Bidfood';
  else if (name.includes('easters')) supplier = 'Easters';
  else if (name.includes('hazels')) supplier = 'Hazels';
  else if (name.includes('makro')) supplier = 'Makro';
  else if (name.includes('pilgrim')) supplier = 'Pilgrim';
  else if (name.includes('broadland')) supplier = 'Broadland';
  else if (name.includes('freestons')) supplier = 'Freestons';
  else notes.push('Supplier could not be inferred from file name.');

  if (name.includes('klm')) {
    site = 'KLM';
  } else {
    notes.push('Site could not be inferred from file name.');
  }

  return {
    supplier: supplier,
    site: site,
    notes: notes.join(' ')
  };
}