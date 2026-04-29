/////////////////////////////////////
// SHEET MECHANICS HELPERS
/////////////////////////////////////

/////////////////////////////////////
// GET TRUE LAST DATA ROW
// VISIBLE DATA ONLY
/////////////////////////////////////
function getLastDataRow(sheet, colLetter) {
  const colIndex = letterToColumn(colLetter);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return 1;

  const values = sheet.getRange(2, colIndex, lastRow - 1, 1).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return i + 2;
    }
  }

  return 1;
}

/////////////////////////////////////
// COLUMN LETTER TO INDEX
/////////////////////////////////////
function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;

  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }

  return column;
}

/////////////////////////////////////
// GET LAST USED ROW IN COLUMN
/////////////////////////////////////
function getLastUsedRowInColumn(sheet, col, startRow) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(startRow, col, maxRows - startRow + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return startRow + i;
    }
  }

  return startRow - 1;
}

function clearInvoiceImportSilent_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Invoice Import');
  if (!sheet) return;

  const startRow = 8;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
  }

  sheet.getRange('B4').clearContent();
  sheet.getRange('B5').clearContent();
}
