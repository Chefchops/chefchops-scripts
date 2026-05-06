/////////////////////////////////////
// PRICE SEARCH
// SEARCH INGREDIENTS MASTER BY TERM
// HEADER-BASED
/////////////////////////////////////

const PRICE_SEARCH_SHEET_NAME_ = 'Price Search';
const PRICE_SEARCH_INPUT_CELL_ = 'B2';


/////////////////////////////////////
// SETUP PRICE SEARCH SHEET
/////////////////////////////////////

function setupPriceSearchSheet() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  let sheet = ss.getSheetByName(PRICE_SEARCH_SHEET_NAME_);

  if (!sheet) {
    sheet = ss.insertSheet(PRICE_SEARCH_SHEET_NAME_);
  }

  sheet.clear();

  sheet.getRange('A1').setValue('Price Search');
  sheet.getRange('A1').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2').setValue('Search Term');
  sheet.getRange('B2').setValue('');

  sheet.getRange('A4').setValue('Examples');
  sheet.getRange('B4').setValue('pickle, fanta, chopped tomatoes, chefs selections');

  const headers = getPriceSearchHeaders_();

  sheet
    .getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.setFrozenRows(6);

  sheet.getRange('A2:B2').setFontWeight('bold');
  sheet.getRange('B2').setBackground('#fff2cc');

  sheet.autoResizeColumns(1, headers.length);

  ui.alert(
    'Price Search sheet setup complete.\n\n' +
    'Type a search term in B2, then run:\n\n' +
    'Chefchops → Price Tools → Run Price Search'
  );
}


/////////////////////////////////////
// RUN PRICE SEARCH
/////////////////////////////////////

function runPriceSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const searchSheet = ss.getSheetByName(PRICE_SEARCH_SHEET_NAME_);
  const masterSheet = ss.getSheetByName('Ingredients Master');

  if (!searchSheet) {
    ui.alert('Price Search sheet not found. Run setupPriceSearchSheet first.');
    return;
  }

  if (!masterSheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const searchTerm = searchSheet
    .getRange(PRICE_SEARCH_INPUT_CELL_)
    .getValue()
    .toString()
    .trim();

  if (!searchTerm) {
    ui.alert('Enter a search term in Price Search cell B2.');
    return;
  }

  clearPriceSearchResults_();

  const masterHeaders = getHeaderMap_(masterSheet, 1);

  const lastRow = masterSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('Ingredients Master has no data.');
    return;
  }

  const requiredHeaders = [
    'Ingredient',
    'Clean Name',
    'Supplier',
    'Pack Size',
    'Pack Price (£)',
    'Cost per Unit (£)'
  ];

  requiredHeaders.forEach(function(headerName) {
    getRequiredHeader_(masterHeaders, headerName, 'Ingredients Master');
  });

  const data = masterSheet
    .getRange(2, 1, lastRow - 1, masterSheet.getLastColumn())
    .getValues();

  const terms = normalisePriceSearchText_(searchTerm)
    .split(' ')
    .filter(Boolean);

  const results = [];

  data.forEach(function(row) {
    const ingredient = getOptionalPriceSearchValue_(row, masterHeaders, 'Ingredient');
    const cleanName = getOptionalPriceSearchValue_(row, masterHeaders, 'Clean Name');
    const category = getOptionalPriceSearchValue_(row, masterHeaders, 'Category');
    const productGroup = getOptionalPriceSearchValue_(row, masterHeaders, 'Product Group');
    const supplier = getOptionalPriceSearchValue_(row, masterHeaders, 'Supplier');
    const packSize = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Size');
    const packQty = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Qty');
    const packPrice = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Price (£)');
    const baseUnit = getOptionalPriceSearchValue_(row, masterHeaders, 'Base Unit');
    const costPerUnit = getOptionalPriceSearchValue_(row, masterHeaders, 'Cost per Unit (£)');
    const itemCode = getOptionalPriceSearchValue_(row, masterHeaders, 'Item Code');
    const notes = getOptionalPriceSearchValue_(row, masterHeaders, 'Notes');

    const searchableText = normalisePriceSearchText_([
      ingredient,
      cleanName,
      category,
      productGroup,
      supplier,
      packSize,
      itemCode,
      notes
    ].join(' '));

    const matchScore = getPriceSearchMatchScore_(searchableText, terms);

    if (matchScore <= 0) return;

    results.push({
      matchScore: matchScore,
      supplier: supplier,
      itemCode: itemCode,
      ingredient: ingredient,
      cleanName: cleanName,
      category: category,
      productGroup: productGroup,
      packSize: packSize,
      packQty: packQty,
      packPrice: packPrice,
      baseUnit: baseUnit,
      costPerUnit: costPerUnit,
      notes: notes
    });
  });

  if (!results.length) {
    ui.alert('No matches found for: ' + searchTerm);
    return;
  }

  results.sort(function(a, b) {
    const scoreDiff = b.matchScore - a.matchScore;
    if (scoreDiff !== 0) return scoreDiff;

    const aCost = parsePriceSearchNumber_(a.costPerUnit);
    const bCost = parsePriceSearchNumber_(b.costPerUnit);

    if (aCost && bCost && aCost !== bCost) return aCost - bCost;

    return String(a.supplier).localeCompare(String(b.supplier));
  });

  const cheapestCost = getCheapestPriceSearchCost_(results);

  const output = results.map(function(item) {
    const cost = parsePriceSearchNumber_(item.costPerUnit);

    const cheapest = cheapestCost !== null && cost === cheapestCost
      ? 'YES'
      : '';

    return [
      cheapest,
      item.matchScore,
      item.supplier,
      item.itemCode,
      item.ingredient,
      item.cleanName,
      item.category,
      item.productGroup,
      item.packSize,
      item.packQty,
      item.packPrice,
      item.baseUnit,
      item.costPerUnit,
      item.notes
    ];
  });

  const headers = getPriceSearchHeaders_();

  searchSheet
    .getRange(7, 1, output.length, headers.length)
    .setValues(output);

  formatPriceSearchResults_(searchSheet, output.length);

  ui.alert(
    'Price Search complete.\n\n' +
    'Search term: ' + searchTerm + '\n' +
    'Matches found: ' + output.length
  );
}


/////////////////////////////////////
// CLEAR PRICE SEARCH RESULTS
/////////////////////////////////////

function clearPriceSearchResults() {
  clearPriceSearchResults_();

  SpreadsheetApp.getUi().alert('Price Search results cleared.');
}


/////////////////////////////////////
// CLEAR PRICE SEARCH RESULTS HELPER
/////////////////////////////////////

function clearPriceSearchResults_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PRICE_SEARCH_SHEET_NAME_);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  if (lastRow > 6) {
    sheet
      .getRange(7, 1, lastRow - 6, sheet.getLastColumn())
      .clearContent()
      .clearFormat();
  }

  const headers = getPriceSearchHeaders_();

  sheet
    .getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.setFrozenRows(6);
}


/////////////////////////////////////
// PRICE SEARCH HEADERS
/////////////////////////////////////

function getPriceSearchHeaders_() {
  return [
    'Cheapest',
    'Match Score',
    'Supplier',
    'Item Code',
    'Ingredient',
    'Clean Name',
    'Category',
    'Product Group',
    'Pack Size',
    'Pack Qty',
    'Pack Price (£)',
    'Base Unit',
    'Cost per Unit (£)',
    'Notes'
  ];
}


/////////////////////////////////////
// FORMAT PRICE SEARCH RESULTS
/////////////////////////////////////

function formatPriceSearchResults_(sheet, resultCount) {
  if (!resultCount) return;

  const headers = getHeaderMap_(sheet, 6);

  const cheapestCol = getRequiredHeader_(headers, 'Cheapest', 'Price Search');
  const packPriceCol = getRequiredHeader_(headers, 'Pack Price (£)', 'Price Search');
  const costPerUnitCol = getRequiredHeader_(headers, 'Cost per Unit (£)', 'Price Search');

  const resultRange = sheet.getRange(7, 1, resultCount, sheet.getLastColumn());

  resultRange.setBorder(true, true, true, true, true, true);

  sheet
    .getRange(7, packPriceCol, resultCount, 1)
    .setNumberFormat('£0.00');

  sheet
    .getRange(7, costPerUnitCol, resultCount, 1)
    .setNumberFormat('£0.0000');

  const cheapestRange = sheet.getRange(7, cheapestCol, resultCount, 1);

  const rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('YES')
      .setBackground('#d9ead3')
      .setRanges([resultRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('YES')
      .setBackground('#93c47d')
      .setRanges([cheapestRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  sheet.autoResizeColumns(1, getPriceSearchHeaders_().length);

  const filterRange = sheet.getRange(6, 1, resultCount + 1, getPriceSearchHeaders_().length);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  filterRange.createFilter();
}


/////////////////////////////////////
// MATCH SCORE
/////////////////////////////////////

function getPriceSearchMatchScore_(searchableText, terms) {
  if (!searchableText || !terms || !terms.length) return 0;

  let score = 0;

  terms.forEach(function(term) {
    if (!term) return;

    if (searchableText === term) {
      score += 100;
    } else if (searchableText.indexOf(term) !== -1) {
      score += 10;
    }
  });

  return score;
}


/////////////////////////////////////
// GET CHEAPEST COST PER UNIT
/////////////////////////////////////

function getCheapestPriceSearchCost_(results) {
  const costs = results
    .map(function(item) {
      return parsePriceSearchNumber_(item.costPerUnit);
    })
    .filter(function(value) {
      return value !== null && value > 0;
    });

  if (!costs.length) return null;

  return Math.min.apply(null, costs);
}


/////////////////////////////////////
// NORMALISE SEARCH TEXT
/////////////////////////////////////

function normalisePriceSearchText_(value) {
  return (value || '')
    .toString()
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/////////////////////////////////////
// PARSE NUMBER / PRICE
/////////////////////////////////////

function parsePriceSearchNumber_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number') return value;

  const cleaned = value
    .toString()
    .replace(/[£,\s]/g, '')
    .trim();

  const number = parseFloat(cleaned);

  return isNaN(number) ? null : number;
}


/////////////////////////////////////
// OPTIONAL VALUE BY HEADER
/////////////////////////////////////

function getOptionalPriceSearchValue_(row, headerMap, headerName) {
  const col = headerMap[headerName];

  if (!col) return '';

  return row[col - 1];
}


/////////////////////////////////////
// RUN PRICE SEARCH
/////////////////////////////////////

function runPriceSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const searchSheet = ss.getSheetByName(PRICE_SEARCH_SHEET_NAME_);
  const masterSheet = ss.getSheetByName('Ingredients Master');

  if (!searchSheet) {
    ui.alert('Price Search sheet not found. Run Setup Price Search first.');
    return;
  }

  if (!masterSheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const searchTerm = searchSheet
    .getRange(PRICE_SEARCH_INPUT_CELL_)
    .getValue()
    .toString()
    .trim();

  if (!searchTerm) {
    ui.alert('Enter a search term in Price Search cell B2.');
    return;
  }

  clearPriceSearchResults_();

  const masterHeaders = getHeaderMap_(masterSheet, 1);

  const requiredHeaders = [
    'Ingredient',
    'Clean Name',
    'Supplier',
    'Pack Size',
    'Pack Price (£)',
    'Cost per Unit (£)'
  ];

  requiredHeaders.forEach(function(headerName) {
    getRequiredHeader_(masterHeaders, headerName, 'Ingredients Master');
  });

  const lastRow = masterSheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('Ingredients Master has no rows to search.');
    return;
  }

  const values = masterSheet
    .getRange(2, 1, lastRow - 1, masterSheet.getLastColumn())
    .getValues();

  const searchTerms = normalisePriceSearchText_(searchTerm)
    .split(' ')
    .filter(Boolean);

  const results = [];

  values.forEach(function(row) {
    const ingredient = getOptionalPriceSearchValue_(row, masterHeaders, 'Ingredient');
    const cleanName = getOptionalPriceSearchValue_(row, masterHeaders, 'Clean Name');
    const category = getOptionalPriceSearchValue_(row, masterHeaders, 'Category');
    const productGroup = getOptionalPriceSearchValue_(row, masterHeaders, 'Product Group');
    const supplier = getOptionalPriceSearchValue_(row, masterHeaders, 'Supplier');
    const packSize = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Size');
    const packQty = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Qty');
    const packPrice = getOptionalPriceSearchValue_(row, masterHeaders, 'Pack Price (£)');
    const baseUnit = getOptionalPriceSearchValue_(row, masterHeaders, 'Base Unit');
    const costPerUnit = getOptionalPriceSearchValue_(row, masterHeaders, 'Cost per Unit (£)');
    const itemCode = getOptionalPriceSearchValue_(row, masterHeaders, 'Item Code');
    const notes = getOptionalPriceSearchValue_(row, masterHeaders, 'Notes');

    const searchableText = normalisePriceSearchText_([
      ingredient,
      cleanName,
      category,
      productGroup,
      supplier,
      packSize,
      itemCode,
      notes
    ].join(' '));

    const matchScore = getPriceSearchMatchScore_(searchableText, searchTerms);

    if (matchScore <= 0) return;

    results.push({
      matchScore: matchScore,
      supplier: supplier,
      itemCode: itemCode,
      ingredient: ingredient,
      cleanName: cleanName,
      category: category,
      productGroup: productGroup,
      packSize: packSize,
      packQty: packQty,
      packPrice: packPrice,
      baseUnit: baseUnit,
      costPerUnit: costPerUnit,
      notes: notes
    });
  });

  if (!results.length) {
    ui.alert('No matches found for: ' + searchTerm);
    return;
  }

  results.sort(function(a, b) {
    const scoreDiff = b.matchScore - a.matchScore;
    if (scoreDiff !== 0) return scoreDiff;

    const aCost = parsePriceSearchNumber_(a.costPerUnit);
    const bCost = parsePriceSearchNumber_(b.costPerUnit);

    if (aCost !== null && bCost !== null && aCost !== bCost) {
      return aCost - bCost;
    }

    return String(a.supplier).localeCompare(String(b.supplier));
  });

  const cheapestCost = getCheapestPriceSearchCost_(results);

  const output = results.map(function(item) {
    const cost = parsePriceSearchNumber_(item.costPerUnit);

    const cheapest = cheapestCost !== null && cost === cheapestCost
      ? 'YES'
      : '';

    return [
      cheapest,
      item.matchScore,
      item.supplier,
      item.itemCode,
      item.ingredient,
      item.cleanName,
      item.category,
      item.productGroup,
      item.packSize,
      item.packQty,
      item.packPrice,
      item.baseUnit,
      item.costPerUnit,
      item.notes
    ];
  });

  const headers = getPriceSearchHeaders_();

  searchSheet
    .getRange(7, 1, output.length, headers.length)
    .setValues(output);

  formatPriceSearchResults_(searchSheet, output.length);

  ui.alert(
    'Price Search complete.\n\n' +
    'Search term: ' + searchTerm + '\n' +
    'Matches found: ' + output.length
  );
}

/////////////////////////////////////
// CLEAR PRICE SEARCH RESULTS
/////////////////////////////////////

function clearPriceSearchResults() {
  clearPriceSearchResults_();

  SpreadsheetApp.getUi().alert('Price Search results cleared.');
}


/////////////////////////////////////
// CLEAR PRICE SEARCH RESULTS HELPER
/////////////////////////////////////

function clearPriceSearchResults_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PRICE_SEARCH_SHEET_NAME_);

  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  if (lastRow > 6) {
    sheet
      .getRange(7, 1, lastRow - 6, sheet.getLastColumn())
      .clearContent()
      .clearFormat();
  }

  const headers = getPriceSearchHeaders_();

  sheet
    .getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.setFrozenRows(6);
}


/////////////////////////////////////
// FORMAT PRICE SEARCH RESULTS
/////////////////////////////////////

function formatPriceSearchResults_(sheet, resultCount) {
  if (!resultCount) return;

  const headers = getHeaderMap_(sheet, 6);

  const cheapestCol = getRequiredHeader_(headers, 'Cheapest', 'Price Search');
  const packPriceCol = getRequiredHeader_(headers, 'Pack Price (£)', 'Price Search');
  const costPerUnitCol = getRequiredHeader_(headers, 'Cost per Unit (£)', 'Price Search');

  const resultRange = sheet.getRange(7, 1, resultCount, getPriceSearchHeaders_().length);

  resultRange.setBorder(true, true, true, true, true, true);

  sheet
    .getRange(7, packPriceCol, resultCount, 1)
    .setNumberFormat('£0.00');

  sheet
    .getRange(7, costPerUnitCol, resultCount, 1)
    .setNumberFormat('£0.0000');

  const cheapestRange = sheet.getRange(7, cheapestCol, resultCount, 1);

  const rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('YES')
      .setBackground('#d9ead3')
      .setRanges([resultRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('YES')
      .setBackground('#93c47d')
      .setRanges([cheapestRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet
    .getRange(6, 1, resultCount + 1, getPriceSearchHeaders_().length)
    .createFilter();

  sheet.autoResizeColumns(1, getPriceSearchHeaders_().length);
}


/////////////////////////////////////
// MATCH SCORE
/////////////////////////////////////

function getPriceSearchMatchScore_(searchableText, terms) {
  if (!searchableText || !terms || !terms.length) return 0;

  let score = 0;

  terms.forEach(function(term) {
    if (!term) return;

    if (searchableText === term) {
      score += 100;
    } else if (searchableText.indexOf(term) !== -1) {
      score += 10;
    }
  });

  return score;
}


/////////////////////////////////////
// GET CHEAPEST COST PER UNIT
/////////////////////////////////////

function getCheapestPriceSearchCost_(results) {
  const costs = results
    .map(function(item) {
      return parsePriceSearchNumber_(item.costPerUnit);
    })
    .filter(function(value) {
      return value !== null && value > 0;
    });

  if (!costs.length) return null;

  return Math.min.apply(null, costs);
}


/////////////////////////////////////
// NORMALISE SEARCH TEXT
/////////////////////////////////////

function normalisePriceSearchText_(value) {
  return (value || '')
    .toString()
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/////////////////////////////////////
// PARSE NUMBER / PRICE
/////////////////////////////////////

function parsePriceSearchNumber_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (typeof value === 'number') return value;

  const cleaned = value
    .toString()
    .replace(/[£,\s]/g, '')
    .trim();

  const number = parseFloat(cleaned);

  return isNaN(number) ? null : number;
}


/////////////////////////////////////
// OPTIONAL VALUE BY HEADER
/////////////////////////////////////

function getOptionalPriceSearchValue_(row, headerMap, headerName) {
  const col = headerMap[headerName];

  if (!col) return '';

  return row[col - 1];
}