/////////////////////////////////////
// ID HELPERS
/////////////////////////////////////

/////////////////////////////////////
// BUILD INGREDIENT ID LOOKUP
/////////////////////////////////////
function buildIngredientIdLookup_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  const headers = getHeaderMap_(sheet, 1);

  const idCol = getRequiredHeader_(headers, 'Ingredient ID', 'Ingredients Master');
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const supplierCol = getRequiredHeader_(headers, 'Supplier', 'Ingredients Master');
  const itemCodeCol = headers['Item Code'];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const lookup = {};

  data.forEach(row => {
    const id = row[idCol - 1];
    const ingredient = (row[ingredientCol - 1] || '').toString().toLowerCase().trim();
    const supplier = (row[supplierCol - 1] || '').toString().toLowerCase().trim();
    const itemCode = itemCodeCol ? (row[itemCodeCol - 1] || '').toString().trim() : '';

    if (!id) return;

    if (itemCode && supplier) {
      lookup[`code|${supplier}|${itemCode}`] = id;
    }

    if (ingredient && supplier) {
      lookup[`name|${supplier}|${ingredient}`] = id;
    }
  });

  return lookup;
}

/////////////////////////////////////
// GENERATE NEXT INGREDIENT ID
/////////////////////////////////////
function generateNextIngredientId_(existingIds) {
  let max = 0;

  existingIds.forEach(id => {
    const match = (id || '').toString().match(/ING(\d+)/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > max) max = num;
    }
  });

  const next = max + 1;
  return 'ING' + Utilities.formatString('%04d', next);
}

/////////////////////////////////////
// ASSIGN MISSING INGREDIENT IDS
/////////////////////////////////////
function assignIngredientIds() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master not found.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);
  const idCol = getRequiredHeader_(headers, 'Ingredient ID', 'Ingredients Master');
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');

  const lastDataRow = getLastDataRowByHeader_(sheet, 1, 'Ingredient');
  if (lastDataRow < 2) {
    ui.alert('No data found.');
    return;
  }

  const data = sheet.getRange(2, 1, lastDataRow - 1, sheet.getLastColumn()).getValues();
  const existingIds = data.map(row => row[idCol - 1]);

  let updated = 0;

  for (let i = 0; i < data.length; i++) {
    const ingredient = (data[i][ingredientCol - 1] || '').toString().trim();
    const id = (data[i][idCol - 1] || '').toString().trim();

    if (!ingredient) continue;

    if (!id) {
      const newId = generateNextIngredientId_(existingIds);
      data[i][idCol - 1] = newId;
      existingIds.push(newId);
      updated++;
    }
  }

  sheet.getRange(2, idCol, data.length, 1)
    .setValues(data.map(row => [row[idCol - 1]]));

  ui.alert(`Ingredient IDs assigned: ${updated}`);
}
