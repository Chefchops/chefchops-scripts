/////////////////////////////////////
// HEADER HELPERS
/////////////////////////////////////

/////////////////////////////////////
// GET HEADER MAP
/////////////////////////////////////
function getHeaderMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const map = {};

  headers.forEach((header, i) => {
    const key = (header || '').toString().trim();
    if (key) map[key] = i + 1; // 1-based
  });

  return map;
}

/////////////////////////////////////
// GET REQUIRED HEADER
/////////////////////////////////////
function getRequiredHeader_(headerMap, headerName, sheetName) {
  const col = headerMap[headerName];
  if (!col) {
    throw new Error(`Missing required header "${headerName}" on sheet "${sheetName}"`);
  }
  return col;
}

/////////////////////////////////////
// GET OPTIONAL HEADER
/////////////////////////////////////
function getOptionalHeader_(headerMap, headerName) {
  return headerMap[headerName] || 0;
}

/////////////////////////////////////
// SET ROW BY HEADERS
/////////////////////////////////////
function setRowByHeaders_(rowArray, headerMap, valuesObj) {
  Object.keys(valuesObj).forEach(key => {
    if (headerMap[key]) {
      rowArray[headerMap[key] - 1] = valuesObj[key];
    }
  });
  return rowArray;
}

/////////////////////////////////////
// GET TRUE LAST DATA ROW BY HEADER
/////////////////////////////////////
function getLastDataRowByHeader_(sheet, headerRow, headerName) {
  const headerMap = getHeaderMap_(sheet, headerRow);
  const colIndex = headerMap[headerName];

  if (!colIndex) {
    throw new Error(`Header "${headerName}" not found on sheet "${sheet.getName()}"`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  const values = sheet.getRange(headerRow + 1, colIndex, lastRow - headerRow, 1).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return headerRow + 1 + i;
    }
  }

  return headerRow;
}

/////////////////////////////////////
// GET LAST REAL DATA ROW BY HEADER
/////////////////////////////////////

function getLastRealDataRowByHeader_(sheet, headerName, headerRow) {
  const headerMap = getHeaderMap_(sheet, headerRow || 1);
  const col = headerMap[headerName];

  if (!col) {
    throw new Error('Header not found: ' + headerName + ' in ' + sheet.getName());
  }

  const startRow = (headerRow || 1) + 1;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) return startRow - 1;

  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return startRow + i;
    }
  }

  return startRow - 1;
}