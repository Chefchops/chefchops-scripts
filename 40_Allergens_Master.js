/////////////////////////////////////
// ALLERGEN HEADER LISTS
/////////////////////////////////////

function getContainsAllergenHeaders_() {
  return [
    'Contains Gluten',
    'Contains Crustaceans',
    'Contains Eggs',
    'Contains Fish',
    'Contains Peanuts',
    'Contains Soya',
    'Contains Milk',
    'Contains Nuts',
    'Contains Celery',
    'Contains Mustard',
    'Contains Sesame',
    'Contains Sulphur Dioxide/Sulphites',
    'Contains Lupin',
    'Contains Molluscs'
  ];
}

function getMayContainAllergenHeaders_() {
  return [
    'May Contain Gluten',
    'May Contain Crustaceans',
    'May Contain Eggs',
    'May Contain Fish',
    'May Contain Peanuts',
    'May Contain Soya',
    'May Contain Milk',
    'May Contain Nuts',
    'May Contain Celery',
    'May Contain Mustard',
    'May Contain Sesame',
    'May Contain Sulphur Dioxide/Sulphites',
    'May Contain Lupin',
    'May Contain Molluscs'
  ];
}

function getAllergenHeaders_() {
  return getContainsAllergenHeaders_().concat(getMayContainAllergenHeaders_());
}

/////////////////////////////////////
// MASTER ALLERGEN SETUP
/////////////////////////////////////

function setupIngredientsMasterAllergenColumns() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const headerRow = 1;
  const headers = getHeaderMap_(sheet, headerRow);
  const allergenHeaders = getAllergenHeaders_();

  /////////////////////////////////////
  // ADD ANY MISSING HEADERS
  /////////////////////////////////////

  const existingHeaderNames = Object.keys(headers);
  const missingHeaders = allergenHeaders.filter(function(header) {
    return !existingHeaderNames.includes(header);
  });

  if (missingHeaders.length > 0) {
    const insertAtCol = getBestAllergenInsertColumn_(sheet, headers);
    sheet.insertColumnsBefore(insertAtCol, missingHeaders.length);
    sheet.getRange(headerRow, insertAtCol, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  /////////////////////////////////////
  // REFRESH HEADER MAP AFTER INSERT
  /////////////////////////////////////

  const updatedHeaders = getHeaderMap_(sheet, headerRow);

  /////////////////////////////////////
  // APPLY CHECKBOXES
  /////////////////////////////////////

  applyIngredientMasterAllergenCheckboxes_(sheet, updatedHeaders);
  formatIngredientsMasterAllergenHeaders_();

  ui.alert(
    'Ingredients Master allergen setup complete.\n\n' +
    'Added missing headers: ' + missingHeaders.length + '\n' +
    'Checkboxes applied to allergen columns.'
  );
}

/////////////////////////////////////
// APPLY ALLERGEN CHECKBOXES
/////////////////////////////////////

function applyIngredientMasterAllergenCheckboxes_(sheet, headerMap) {
  const startRow = 2;
  const allergenHeaders = getAllergenHeaders_();

  const ingredientCol = getOptionalHeader_(headerMap, 'Ingredient');
  const lastDataRow = ingredientCol
    ? getLastRealDataRowInColumn_(sheet, ingredientCol, startRow)
    : sheet.getLastRow();

  const bufferRows = 50;
  const endRow = Math.max(startRow, lastDataRow + bufferRows);
  const numRows = endRow - startRow + 1;

  if (numRows < 1) return;

  allergenHeaders.forEach(function(header) {
    const col = getOptionalHeader_(headerMap, header);
    if (!col) return;

    const range = sheet.getRange(startRow, col, numRows, 1);
    const values = range.getValues();

    const cleaned = values.map(function(row) {
      const value = row[0];

      if (value === true || value === 'TRUE') return [true];
      if (value === false || value === 'FALSE') return [false];
      return [''];
    });

    range.clearDataValidations();
    range.removeCheckboxes();
    range.insertCheckboxes();
    range.setValues(cleaned);
  });
}

/////////////////////////////////////
// BEST INSERT POSITION
/////////////////////////////////////

function getBestAllergenInsertColumn_(sheet, headerMap) {
  const preferredBeforeHeaders = [
    'Pack Size',
    'Pack Qty',
    'Pack Price (£)',
    'Pack Price',
    'Base Unit',
    'Cost per Unit (£)',
    'Cost per Unit',
    'Supplier'
  ];

  for (let i = 0; i < preferredBeforeHeaders.length; i++) {
    const col = getOptionalHeader_(headerMap, preferredBeforeHeaders[i]);
    if (col) return col;
  }

  return sheet.getLastColumn() + 1;
}

/////////////////////////////////////
// FORMAT ALLERGEN HEADERS
/////////////////////////////////////

function formatIngredientsMasterAllergenHeaders_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  if (!sheet) return;

  const headerRow = 1;
  const headerMap = getHeaderMap_(sheet, headerRow);

  const containsHeaders = getContainsAllergenHeaders_();
  const mayContainHeaders = getMayContainAllergenHeaders_();

  containsHeaders.forEach(function(header) {
    const col = getOptionalHeader_(headerMap, header);
    if (!col) return;

    sheet.getRange(headerRow, col)
      .setBackground('#ea9999')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
  });

  mayContainHeaders.forEach(function(header) {
    const col = getOptionalHeader_(headerMap, header);
    if (!col) return;

    sheet.getRange(headerRow, col)
      .setBackground('#f9cb9c')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setWrap(true);
  });
}

/////////////////////////////////////
// OPTIONAL COLOUR FORMATTING
/////////////////////////////////////

function formatIngredientsMasterAllergenColumns() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const headerRow = 1;
  const headerMap = getHeaderMap_(sheet, headerRow);
  const containsHeaders = getContainsAllergenHeaders_();
  const mayContainHeaders = getMayContainAllergenHeaders_();

  const startRow = 2;
  const ingredientCol = getOptionalHeader_(headerMap, 'Ingredient');
  const lastDataRow = ingredientCol
    ? getLastRealDataRowInColumn_(sheet, ingredientCol, startRow)
    : sheet.getLastRow();

  const bufferRows = 50;
  const endRow = Math.max(startRow, lastDataRow + bufferRows);
  const numRows = endRow - startRow + 1;

  if (numRows < 1) {
    ui.alert('No allergen rows available to format.');
    return;
  }

  const existingRules = sheet.getConditionalFormatRules() || [];
  const keptRules = existingRules.filter(function(rule) {
    const ranges = rule.getRanges() || [];

    for (let i = 0; i < ranges.length; i++) {
      const range = ranges[i];
      const row = range.getRow();
      const numCols = range.getNumColumns();

      if (row === startRow && numCols === 1) {
        const col = range.getColumn();
        const header = Object.keys(headerMap).find(function(name) {
          return headerMap[name] === col;
        });

        if (
          containsHeaders.includes(header) ||
          mayContainHeaders.includes(header)
        ) {
          return false;
        }
      }
    }

    return true;
  });

  const newRules = [];

  containsHeaders.forEach(function(header) {
    const col = getOptionalHeader_(headerMap, header);
    if (!col) return;

    const range = sheet.getRange(startRow, col, numRows, 1);

    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + columnToLetter_(col) + startRow + '=TRUE')
        .setBackground('#f4cccc')
        .setRanges([range])
        .build()
    );
  });

  mayContainHeaders.forEach(function(header) {
    const col = getOptionalHeader_(headerMap, header);
    if (!col) return;

    const range = sheet.getRange(startRow, col, numRows, 1);

    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + columnToLetter_(col) + startRow + '=TRUE')
        .setBackground('#fce5cd')
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(keptRules.concat(newRules));

  ui.alert('Allergen colour formatting applied.');
}

/////////////////////////////////////
// GET LAST REAL DATA ROW IN COLUMN
/////////////////////////////////////

function getLastRealDataRowInColumn_(sheet, columnNumber, startRow) {
  const firstRow = startRow || 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) return firstRow - 1;

  const values = sheet.getRange(firstRow, columnNumber, lastRow - firstRow + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return firstRow + i;
    }
  }

  return firstRow - 1;
}

/////////////////////////////////////
// HELPER
/////////////////////////////////////

function columnToLetter_(column) {
  let temp = '';
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}