/////////////////////////////////////
// BUILD SUPPLIER COMPARISON
/////////////////////////////////////

function buildSupplierComparison() {
  const t0 = new Date().getTime();

  function logStep_(label) {
    const now = new Date().getTime();
    Logger.log(label + ': ' + (now - t0) + ' ms');
  }

  const ss = SpreadsheetApp.getActive();
  const sourceSheet = ss.getSheetByName('Ingredients Master');
  const compareSheet = ss.getSheetByName('Price Comparison');
  const ui = SpreadsheetApp.getUi();

  logStep_('Start');

  if (!sourceSheet || !compareSheet) {
    ui.alert('Missing required sheet: Ingredients Master or Price Comparison.');
    return;
  }

  const sourceHeaderMap = getHeaderMap_(sourceSheet, 1);

  const colIngredientId = getRequiredHeader_(sourceHeaderMap, 'Ingredient ID', 'Ingredients Master');
  const colIngredient = getRequiredHeader_(sourceHeaderMap, 'Ingredient', 'Ingredients Master');
  const colCleanName = getRequiredHeader_(sourceHeaderMap, 'Clean Name', 'Ingredients Master');
  const colSupplier = getRequiredHeader_(sourceHeaderMap, 'Supplier', 'Ingredients Master');
  const colPackSize = getRequiredHeader_(sourceHeaderMap, 'Pack Size', 'Ingredients Master');
  const colPackQty = getRequiredHeader_(sourceHeaderMap, 'Pack Qty', 'Ingredients Master');
  const colPackPrice = getRequiredHeader_(sourceHeaderMap, 'Pack Price (£)', 'Ingredients Master');
  const colBaseUnit = getRequiredHeader_(sourceHeaderMap, 'Base Unit', 'Ingredients Master');
  const colCostPerUnit = getRequiredHeader_(sourceHeaderMap, 'Cost per Unit (£)', 'Ingredients Master');
  const colItemCode = getRequiredHeader_(sourceHeaderMap, 'Item Code', 'Ingredients Master');

  logStep_('Headers resolved');

  const searchValue = (compareSheet.getRange('B3').getValue() || '').toString().trim();
  if (!searchValue) {
    ui.alert('Enter an ingredient search in cell B3.');
    return;
  }

  logStep_('Search value read');

  const lastRow = getLastRealDataRowByHeader_(sourceSheet, 'Ingredient', 1);
  if (lastRow < 2) {
    ui.alert('Ingredients Master has no data.');
    return;
  }

  logStep_('Last real row found');

  const maxCol = Math.max(
    colIngredientId,
    colIngredient,
    colCleanName,
    colSupplier,
    colPackSize,
    colPackQty,
    colPackPrice,
    colBaseUnit,
    colCostPerUnit,
    colItemCode
  );

  const data = sourceSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  const searchLower = searchValue.toLowerCase();

  logStep_('Source data read');

  const latestMap = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    const ingredient = (row[colIngredient - 1] || '').toString().trim();
    const cleanName = (row[colCleanName - 1] || '').toString().trim();

    const ingredientLower = ingredient.toLowerCase();
    const cleanNameLower = cleanName.toLowerCase();

    if (ingredientLower.indexOf(searchLower) === -1 && cleanNameLower.indexOf(searchLower) === -1) {
      continue;
    }

    const ingredientId = (row[colIngredientId - 1] || '').toString().trim();
    const supplier = (row[colSupplier - 1] || '').toString().trim();
    const packSize = (row[colPackSize - 1] || '').toString().trim();
    const packQty = row[colPackQty - 1];
    const packPrice = toNumber(row[colPackPrice - 1]);
    const baseUnit = (row[colBaseUnit - 1] || '').toString().trim();
    const costPerUnit = toNumber(row[colCostPerUnit - 1]);
    const itemCode = (row[colItemCode - 1] || '').toString().trim();

    const key = [
      cleanNameLower,
      supplier.toLowerCase(),
      packSize.toLowerCase(),
      String(packQty).toLowerCase(),
      baseUnit.toLowerCase()
    ].join('|');

    latestMap[key] = [
      ingredientId,
      ingredient,
      cleanName,
      supplier,
      packSize,
      packQty,
      baseUnit,
      isNaN(packPrice) ? '' : packPrice,
      isNaN(costPerUnit) ? '' : costPerUnit,
      itemCode
    ];
  }

  logStep_('Filtering and map build done');

  const output = Object.values(latestMap);

  output.sort((a, b) => {
    const aVal = toNumber(a[8]);
    const bVal = toNumber(b[8]);

    if (isNaN(aVal) && isNaN(bVal)) return 0;
    if (isNaN(aVal)) return 1;
    if (isNaN(bVal)) return -1;
    return aVal - bVal;
  });

  logStep_('Output sorted');

  const clearRows = Math.max(compareSheet.getMaxRows() - 6, 1);
  compareSheet.getRange(6, 2, clearRows, 10).clearContent().clearFormat();

  logStep_('Output area cleared');

  const headers = [[
    'Ingredient ID',
    'Ingredient',
    'Clean Name',
    'Supplier',
    'Pack Size',
    'Pack Qty',
    'Base Unit',
    'Pack Price (£)',
    'Cost per Unit (£)',
    'Item Code'
  ]];

  compareSheet.getRange(6, 2, 1, headers[0].length).setValues(headers);
  compareSheet.getRange(6, 2, 1, headers[0].length).setFontWeight('bold');

  logStep_('Headers written');

  if (output.length === 0) {
    compareSheet.getRange('B7').setValue('No matches found.');
    logStep_('Finished - no matches');
    ui.alert('No matches found.');
    return;
  }

  compareSheet.getRange(7, 2, output.length, output[0].length).setValues(output);
  compareSheet.getRange(7, 9, output.length, 1).setNumberFormat('£0.0000');
  compareSheet.getRange(7, 10, output.length, 1).setNumberFormat('£0.0000');

  logStep_('Results written');

  let cheapestRowIndex = -1;
  let cheapestValue = null;

  for (let i = 0; i < output.length; i++) {
    const cpu = toNumber(output[i][8]);
    if (isNaN(cpu)) continue;

    if (cheapestValue === null || cpu < cheapestValue) {
      cheapestValue = cpu;
      cheapestRowIndex = i;
    }
  }

  logStep_('Cheapest row calculated');

  if (cheapestRowIndex !== -1) {
    compareSheet.getRange(7 + cheapestRowIndex, 2, 1, 10).setBackground('#d9ead3');
  }

  logStep_('Cheapest row highlighted');
  logStep_('Finished');

  ui.alert('Price comparison built.');
}

////////////////////////////////
// OLD PRICE COMP LEGACY DO NOT USE
////////////////////////////////////////

function OLD_priceComparisonSearchPopup() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Price Comparison');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Missing "Price Comparison" sheet.');
    return;
  }

  const response = ui.prompt(
    'Price Comparison Search',
    'Enter ingredient to search:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const searchText = (response.getResponseText() || '').toString().trim();

  if (!searchText) {
    ui.alert('No search entered.');
    return;
  }

  ss.setActiveSheet(sheet);
  sheet.getRange('B1').setValue(searchText);
  sheet.getRange('A4').activate();
}



/////////////////////////////////////
// PRICE SEARCH UI
/////////////////////////////////////

function openBestPriceSearch() {
  const html = HtmlService.createHtmlOutputFromFile('Price_Search_UI')
    .setWidth(900)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'Best Price Search');
}

/////////////////////////////////////
// SEARCH PRICE RESULTS
/////////////////////////////////////

function searchBestPrices(searchText) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  if (!sheet) {
    throw new Error('Ingredients Master sheet not found.');
  }

  const headerRow = 1;
  const headers = getHeaderMap_(sheet, headerRow);

  const colIngredientId = getRequiredHeader_(headers, 'Ingredient ID', 'Ingredients Master');
  const colIngredient = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const colCleanName = getRequiredHeader_(headers, 'Clean Name', 'Ingredients Master');
  const colSupplier = getRequiredHeader_(headers, 'Supplier', 'Ingredients Master');
  const colPackSize = getRequiredHeader_(headers, 'Pack Size', 'Ingredients Master');
  const colPackQty = getRequiredHeader_(headers, 'Pack Qty', 'Ingredients Master');
  const colBaseUnit = getRequiredHeader_(headers, 'Base Unit', 'Ingredients Master');
  const colPackPrice = getRequiredHeader_(headers, 'Pack Price (£)', 'Ingredients Master');
  const colCostPerUnit = getRequiredHeader_(headers, 'Cost per Unit (£)', 'Ingredients Master');
  const colItemCode = getOptionalHeader_(headers, 'Item Code');

  const lastRow = getLastDataRowByHeader_(sheet, headerRow, 'Ingredient');
  if (lastRow <= headerRow) return [];

  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

  const term = (searchText || '').toString().trim().toLowerCase();
  if (!term) return [];

  const results = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const ingredientId = row[colIngredientId - 1];
    const ingredient = row[colIngredient - 1];
    const cleanName = row[colCleanName - 1];
    const supplier = row[colSupplier - 1];
    const packSize = row[colPackSize - 1];
    const packQty = row[colPackQty - 1];
    const baseUnit = row[colBaseUnit - 1];
    const packPrice = row[colPackPrice - 1];
    const costPerUnit = row[colCostPerUnit - 1];
    const itemCode = colItemCode ? row[colItemCode - 1] : '';

    if (!ingredient || !supplier || costPerUnit === '' || costPerUnit === null) continue;

    const ingredientText = (ingredient || '').toString().toLowerCase();
    const cleanText = (cleanName || '').toString().toLowerCase();

    if (ingredientText.indexOf(term) === -1 && cleanText.indexOf(term) === -1) continue;

    results.push({
      ingredientId: ingredientId || '',
      ingredient: ingredient || '',
      cleanName: cleanName || '',
      supplier: supplier || '',
      packSize: packSize || '',
      packQty: packQty || '',
      baseUnit: baseUnit || '',
      packPrice: Number(packPrice) || 0,
      costPerUnit: Number(costPerUnit) || 0,
      itemCode: itemCode || ''
    });
  }

  results.sort(function(a, b) {
    return a.costPerUnit - b.costPerUnit;
  });

  if (results.length > 0) {
    const best = results[0].costPerUnit;
    for (let i = 0; i < results.length; i++) {
      results[i].isBest = results[i].costPerUnit === best;
    }
  }

  return results;
}