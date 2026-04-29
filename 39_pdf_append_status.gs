/////////////////////////////////////
// ENSURE PDF STAGING APPEND STATUS
/////////////////////////////////////

function ensurePdfStagingAppendStatusColumn_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');

  if (!sheet) throw new Error('Missing sheet: PDF Staging');

  const headers = getHeaderMap_(sheet, 1);

  if (headers['Append Status']) return headers['Append Status'];

  const newCol = sheet.getLastColumn() + 1;

  sheet
    .getRange(1, newCol)
    .setValue('Append Status')
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.autoResizeColumn(newCol);

  return newCol;
}

/////////////////////////////////////
// GET APPEND STATUS FOR FILE
/////////////////////////////////////

function getPdfAppendStatusForFile_(fileId) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');

  if (!sheet) throw new Error('Missing sheet: PDF Staging');

  const appendStatusCol = ensurePdfStagingAppendStatusColumn_();
  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowFileId = values[i][fileIdCol - 1];

    if ((rowFileId || '').toString().trim() === fileId.toString().trim()) {
      return (values[i][appendStatusCol - 1] || '').toString().trim();
    }
  }

  return '';
}

/////////////////////////////////////
// SET APPEND STATUS FOR FILE
/////////////////////////////////////

function setPdfAppendStatusForFile_(fileId, status) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Staging');

  if (!sheet) throw new Error('Missing sheet: PDF Staging');

  const appendStatusCol = ensurePdfStagingAppendStatusColumn_();
  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Staging');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const fileIds = sheet
    .getRange(2, fileIdCol, lastRow - 1, 1)
    .getValues()
    .flat();

  for (let i = 0; i < fileIds.length; i++) {
    if ((fileIds[i] || '').toString().trim() === fileId.toString().trim()) {
      sheet.getRange(i + 2, appendStatusCol).setValue(status);
      return true;
    }
  }

  return false;
}

/////////////////////////////////////
// TEST SETUP APPEND STATUS
/////////////////////////////////////

function setupPdfAppendStatusColumn() {
  ensurePdfStagingAppendStatusColumn_();
  SpreadsheetApp.getUi().alert('PDF Staging Append Status column is ready.');
}