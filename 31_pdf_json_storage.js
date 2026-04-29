/////////////////////////////////////
// PDF JSON STORAGE
// Stores cloud/few-shot JSON safely in chunks
// Header-based + batch delete + fast rebuild
/////////////////////////////////////

const PDF_JSON_STAGING_SHEET_NAME_ = 'PDF JSON Staging';
const PDF_JSON_STAGING_HEADER_ROW_ = 1;
const PDF_JSON_MAX_CHUNK_SIZE_ = 40000;

/////////////////////////////////////
// PDF JSON STAGING HEADERS
/////////////////////////////////////

function getPdfJsonStagingHeaders_() {
  return [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'Chunk No',
    'Chunk Count',
    'JSON Chunk'
  ];
}

/////////////////////////////////////
// SETUP PDF JSON STAGING SHEET
/////////////////////////////////////

function setupPdfJsonStagingSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(PDF_JSON_STAGING_SHEET_NAME_);

  if (!sheet) {
    sheet = ss.insertSheet(PDF_JSON_STAGING_SHEET_NAME_);
  }

  const headers = getPdfJsonStagingHeaders_();

  sheet
    .getRange(PDF_JSON_STAGING_HEADER_ROW_, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sheet.setFrozenRows(1);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet
    .getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length)
    .createFilter();

  sheet.autoResizeColumns(1, headers.length);

  return sheet;
}

/////////////////////////////////////
// GET PDF JSON STAGING SHEET
/////////////////////////////////////

function getPdfJsonStagingSheet_() {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(PDF_JSON_STAGING_SHEET_NAME_);

  if (!sheet) {
    return setupPdfJsonStagingSheet();
  }

  return sheet;
}

/////////////////////////////////////
// WRITE JSON TO STAGING
/////////////////////////////////////

function writeJsonToStaging_(sheet, obj) {
  sheet = sheet || getPdfJsonStagingSheet_();

  if (!obj) throw new Error('Missing JSON staging object.');
  if (!obj.fileId) throw new Error('Missing fileId.');
  if (!obj.jsonText) throw new Error('Missing jsonText.');

  setupPdfJsonStagingSheet();

  clearJsonChunksForFile_(sheet, obj.fileId);

  const jsonBase64 = Utilities.base64Encode(
    Utilities.newBlob(obj.jsonText, 'application/json').getBytes()
  );

  const chunks = [];

  for (let i = 0; i < jsonBase64.length; i += PDF_JSON_MAX_CHUNK_SIZE_) {
    chunks.push(jsonBase64.substring(i, i + PDF_JSON_MAX_CHUNK_SIZE_));
  }

  const startRow = Math.max(sheet.getLastRow() + 1, 2);

  const rows = chunks.map((chunk, index) => [
    obj.uploadTime || new Date(),
    obj.fileName || '',
    obj.supplier || '',
    obj.site || '',
    obj.fileId,
    index + 1,
    chunks.length,
    chunk
  ]);

  sheet
    .getRange(startRow, 1, rows.length, rows[0].length)
    .setValues(rows);
}

/////////////////////////////////////
// CLEAR JSON CHUNKS FOR FILE
// Batch delete, bottom-up grouped ranges
/////////////////////////////////////

function clearJsonChunksForFile_(sheet, fileId) {
  sheet = sheet || getPdfJsonStagingSheet_();

  if (!fileId) throw new Error('Missing fileId.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headerMap = getHeaderMap_(sheet, PDF_JSON_STAGING_HEADER_ROW_);
  const fileIdCol = getRequiredHeader_(
    headerMap,
    'Drive File ID',
    PDF_JSON_STAGING_SHEET_NAME_
  );

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
// DELETE ROWS IN GROUPS
// Much faster than deleteRow repeatedly
/////////////////////////////////////
function deleteRowsInGroups_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;

  rowNumbers = rowNumbers
    .map(Number)
    .filter(row => row > 1)
    .sort((a, b) => b - a);

  if (!rowNumbers.length) return;

  const maxRows = sheet.getMaxRows();

  // If deleting all non-header rows, clear contents instead of deleting rows
  if (rowNumbers.length >= maxRows - 1) {
    sheet
      .getRange(2, 1, maxRows - 1, sheet.getLastColumn())
      .clearContent();

    return;
  }

  let groupStart = rowNumbers[0];
  let groupCount = 1;

  for (let i = 1; i < rowNumbers.length; i++) {
    const row = rowNumbers[i];

    if (row === groupStart - groupCount) {
      groupCount++;
    } else {
      sheet.deleteRows(groupStart - groupCount + 1, groupCount);
      groupStart = row;
      groupCount = 1;
    }
  }

  sheet.deleteRows(groupStart - groupCount + 1, groupCount);
}
/////////////////////////////////////
// REBUILD JSON FROM CHUNKS
/////////////////////////////////////

function rebuildJsonFromChunks_(fileId) {
  const sheet = getPdfJsonStagingSheet_();

  if (!fileId) throw new Error('Missing fileId.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('No JSON chunk data found.');
  }

  const headerMap = getHeaderMap_(sheet, PDF_JSON_STAGING_HEADER_ROW_);

  const fileIdCol = getRequiredHeader_(headerMap, 'Drive File ID', PDF_JSON_STAGING_SHEET_NAME_);
  const chunkNoCol = getRequiredHeader_(headerMap, 'Chunk No', PDF_JSON_STAGING_SHEET_NAME_);
  const chunkCountCol = getRequiredHeader_(headerMap, 'Chunk Count', PDF_JSON_STAGING_SHEET_NAME_);
  const jsonChunkCol = getRequiredHeader_(headerMap, 'JSON Chunk', PDF_JSON_STAGING_SHEET_NAME_);

  const maxCol = Math.max(fileIdCol, chunkNoCol, chunkCountCol, jsonChunkCol);

  const data = sheet
    .getRange(2, 1, lastRow - 1, maxCol)
    .getValues();

  const rows = data
    .filter(row => {
      return (row[fileIdCol - 1] || '').toString().trim() === fileId.toString().trim();
    })
    .sort((a, b) => {
      return Number(a[chunkNoCol - 1]) - Number(b[chunkNoCol - 1]);
    });

  if (!rows.length) {
    throw new Error('No JSON chunks found for file ID: ' + fileId);
  }

  const expectedChunkCount = Number(rows[0][chunkCountCol - 1]);

  if (rows.length !== expectedChunkCount) {
    throw new Error(
      'JSON chunk count mismatch for file ID: ' +
      fileId +
      '. Expected ' +
      expectedChunkCount +
      ', found ' +
      rows.length +
      '.'
    );
  }

  const base64Text = rows
    .map(row => (row[jsonChunkCol - 1] || '').toString())
    .join('');

  const jsonText = Utilities
    .newBlob(Utilities.base64Decode(base64Text), 'application/json')
    .getDataAsString();

  return JSON.parse(jsonText);
}

/////////////////////////////////////
// GET PDF JSON META BY FILE ID
/////////////////////////////////////

function getPdfJsonMetaByFileId_(fileId) {
  const sheet = getPdfJsonStagingSheet_();

  if (!fileId) throw new Error('Missing fileId.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const headerMap = getHeaderMap_(sheet, PDF_JSON_STAGING_HEADER_ROW_);

  const uploadTimeCol = getRequiredHeader_(headerMap, 'Upload Time', PDF_JSON_STAGING_SHEET_NAME_);
  const fileNameCol = getRequiredHeader_(headerMap, 'File Name', PDF_JSON_STAGING_SHEET_NAME_);
  const supplierCol = getRequiredHeader_(headerMap, 'Supplier', PDF_JSON_STAGING_SHEET_NAME_);
  const siteCol = getRequiredHeader_(headerMap, 'Site', PDF_JSON_STAGING_SHEET_NAME_);
  const fileIdCol = getRequiredHeader_(headerMap, 'Drive File ID', PDF_JSON_STAGING_SHEET_NAME_);

  const maxCol = Math.max(uploadTimeCol, fileNameCol, supplierCol, siteCol, fileIdCol);

  const data = sheet
    .getRange(2, 1, lastRow - 1, maxCol)
    .getValues();

  const row = data.find(r => {
    return (r[fileIdCol - 1] || '').toString().trim() === fileId.toString().trim();
  });

  if (!row) return null;

  return {
    uploadTime: row[uploadTimeCol - 1],
    fileName: row[fileNameCol - 1],
    supplier: row[supplierCol - 1],
    site: row[siteCol - 1],
    fileId: row[fileIdCol - 1]
  };
}

/////////////////////////////////////
// TEST PDF JSON REBUILD
/////////////////////////////////////

function testRebuildJsonFromChunks() {
  const fileId = Browser.inputBox('Enter Drive File ID to rebuild JSON');

  if (!fileId || fileId === 'cancel') return;

  const json = rebuildJsonFromChunks_(fileId);

  SpreadsheetApp.getUi().alert(
    'JSON rebuilt successfully.\n\nTop-level keys:\n' +
    Object.keys(json).join('\n')
  );
}