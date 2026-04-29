/////////////////////////////////////
// CATEGORY + PRODUCT GROUP HELPERS
/////////////////////////////////////

/////////////////////////////////////
// CHECK MISSING PRODUCT GROUPS
/////////////////////////////////////
function checkMissingProductGroups() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master not found.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);
  const colIngredient = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const colProductGroup = getRequiredHeader_(headers, 'Product Group', 'Ingredients Master');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data found.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const missing = [];

  for (let i = 0; i < data.length; i++) {
    const ingredient = (data[i][colIngredient - 1] || '').toString().trim();
    const productGroup = (data[i][colProductGroup - 1] || '').toString().trim();

    if (!productGroup && ingredient) {
      missing.push(ingredient);
    }
  }

  if (missing.length === 0) {
    ui.alert('All items have product groups ✅');
    return;
  }

  let message = `Missing Product Groups: ${missing.length}\n\n`;

  missing.slice(0, 10).forEach(item => {
    message += `• ${item}\n`;
  });

  if (missing.length > 10) {
    message += `\n...and ${missing.length - 10} more`;
  }

  ui.alert(message);
}

/////////////////////////////////////
// AUTO-SUGGEST PRODUCT GROUPS
// Fills blank Product Group cells only
/////////////////////////////////////
function autoSuggestIngredientProductGroups() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No ingredient rows found.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const productGroupCol = getRequiredHeader_(headers, 'Product Group', 'Ingredients Master');
  const cleanNameCol = headers['Clean Name'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const existingLookup = buildExistingProductGroupLookup_();
  const rules = getProductGroupRules_();

  let updated = 0;
  const output = [];

  for (let i = 0; i < data.length; i++) {
    const ingredient = (data[i][ingredientCol - 1] || '').toString().trim();
    const cleanName = cleanNameCol ? (data[i][cleanNameCol - 1] || '').toString().trim() : '';
    const existingProductGroup = (data[i][productGroupCol - 1] || '').toString().trim();

    if (existingProductGroup) {
      output.push([existingProductGroup]);
      continue;
    }

    let suggestedProductGroup = '';

    const ingredientKey = ingredient.toLowerCase();
    const cleanNameKey = cleanName.toLowerCase();
    const combinedText = `${ingredient} ${cleanName}`.trim().toLowerCase();

    if (ingredientKey && existingLookup.byIngredient[ingredientKey]) {
      suggestedProductGroup = existingLookup.byIngredient[ingredientKey];
    } else {
      for (const rule of rules) {
        if (matchesProductGroupRule_(combinedText, rule)) {
          suggestedProductGroup = rule.productGroup;
          break;
        }
      }

      if (!suggestedProductGroup && cleanNameKey && existingLookup.byCleanName[cleanNameKey]) {
        suggestedProductGroup = existingLookup.byCleanName[cleanNameKey];
      }
    }

    if (suggestedProductGroup) updated++;
    output.push([suggestedProductGroup]);
  }

  sheet.getRange(2, productGroupCol, output.length, 1).setValues(output);

  ui.alert(`Product Group auto-suggest complete.\n\nUpdated blank product group cells: ${updated}`);
}

/////////////////////////////////////
// PRODUCT GROUP MATCH HELPER
/////////////////////////////////////
function matchesProductGroupRule_(text, rule) {
  if (!text || !rule || !rule.matchText) return false;

  const source = text.toLowerCase().trim();
  const target = rule.matchText.toLowerCase().trim();

  switch (rule.matchType) {
    case 'contains':
      return source.indexOf(target) !== -1;
    case 'equals':
      return source === target;
    case 'starts_with':
      return source.startsWith(target);
    case 'ends_with':
      return source.endsWith(target);
    default:
      return false;
  }
}

/////////////////////////////////////
// EXISTING PRODUCT GROUP LOOKUP
/////////////////////////////////////
function buildExistingProductGroupLookup_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  if (!sheet) {
    throw new Error('Ingredients Master not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      byIngredient: {},
      byCleanName: {}
    };
  }

  const headers = getHeaderMap_(sheet, 1);
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const productGroupCol = getRequiredHeader_(headers, 'Product Group', 'Ingredients Master');
  const cleanNameCol = headers['Clean Name'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const byIngredient = {};
  const byCleanName = {};

  data.forEach(row => {
    const ingredient = (row[ingredientCol - 1] || '').toString().trim().toLowerCase();
    const productGroup = (row[productGroupCol - 1] || '').toString().trim();
    const cleanName = cleanNameCol ? (row[cleanNameCol - 1] || '').toString().trim().toLowerCase() : '';

    if (ingredient && productGroup && !byIngredient[ingredient]) {
      byIngredient[ingredient] = productGroup;
    }

    if (cleanName && productGroup && !byCleanName[cleanName]) {
      byCleanName[cleanName] = productGroup;
    }
  });

  return { byIngredient, byCleanName };
}

/////////////////////////////////////
// PRODUCT GROUP RULES FROM SHEET
/////////////////////////////////////
function getProductGroupRules_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Product Group Rules');

  if (!sheet) {
    throw new Error('Missing "Product Group Rules" sheet.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = getHeaderMap_(sheet, 1);

  const productGroupCol = getRequiredHeader_(headers, 'Product Group', 'Product Group Rules');
  const matchTextCol = getRequiredHeader_(headers, 'Match Text', 'Product Group Rules');
  const matchTypeCol = getRequiredHeader_(headers, 'Match Type', 'Product Group Rules');
  const priorityCol = headers['Priority'];
  const enabledCol = headers['Enabled'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data
    .map(row => {
      const productGroup = (row[productGroupCol - 1] || '').toString().trim();
      const matchText = (row[matchTextCol - 1] || '').toString().trim().toLowerCase();
      const matchType = (row[matchTypeCol - 1] || '').toString().trim().toLowerCase();
      const priority = priorityCol ? Number(row[priorityCol - 1]) || 999 : 999;
      const enabled = enabledCol ? (row[enabledCol - 1] || '').toString().trim().toUpperCase() : 'Y';

      return {
        productGroup,
        matchText,
        matchType,
        priority,
        enabled
      };
    })
    .filter(r => r.productGroup && r.matchText && r.matchType && r.enabled !== 'N')
    .sort((a, b) => a.priority - b.priority);
}

/////////////////////////////////////
// INGREDIENT CATEGORY DROPDOWN
/////////////////////////////////////
function setIngredientCategoryDropdown() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Ingredients Master');
  const lists = ss.getSheetByName('Lists');
  const ui = SpreadsheetApp.getUi();

  if (!master || !lists) {
    ui.alert('Missing Ingredients Master or Lists sheet.');
    return;
  }

  const masterHeaders = getHeaderMap_(master, 1);
  const listsHeaders = getHeaderMap_(lists, 1);

  const masterCategoryCol = getRequiredHeader_(masterHeaders, 'Category', 'Ingredients Master');

  let listsCategoryCol = listsHeaders['Category'];
  if (!listsCategoryCol) {
    listsCategoryCol = 1;
  }

  const lastRow = lists.getLastRow();
  if (lastRow < 2) {
    ui.alert('No categories found on Lists sheet.');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lists.getRange(2, listsCategoryCol, lastRow - 1, 1), true)
    .setAllowInvalid(false)
    .build();

  master
    .getRange(2, masterCategoryCol, master.getMaxRows() - 1, 1)
    .clearDataValidations()
    .setDataValidation(rule);

  ui.alert('Category dropdown applied to Ingredients Master Category column.');
}

/////////////////////////////////////
// PRODUCT GROUP DROPDOWN FROM LISTS
/////////////////////////////////////
function setIngredientProductGroupDropdown() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Ingredients Master');
  const lists = ss.getSheetByName('Lists');
  const ui = SpreadsheetApp.getUi();

  if (!master || !lists) {
    ui.alert('Missing Ingredients Master or Lists sheet.');
    return;
  }

  const masterHeaders = getHeaderMap_(master, 1);
  const listsHeaders = getHeaderMap_(lists, 1);

  const masterProductGroupCol = getRequiredHeader_(masterHeaders, 'Product Group', 'Ingredients Master');

  let listsProductGroupCol = listsHeaders['Product Group'];
  if (!listsProductGroupCol) {
    listsProductGroupCol = 2;
  }

  const lastRow = lists.getLastRow();
  if (lastRow < 2) {
    ui.alert('No product groups found on Lists sheet.');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lists.getRange(2, listsProductGroupCol, lastRow - 1, 1), true)
    .setAllowInvalid(true)
    .build();

  master
    .getRange(2, masterProductGroupCol, master.getMaxRows() - 1, 1)
    .clearDataValidations()
    .setDataValidation(rule);

  ui.alert('Product Group dropdown applied to Ingredients Master Product Group column.');
}

/////////////////////////////////////
// CHECK UNASSIGNED CATEGORIES
/////////////////////////////////////
function checkMissingCategories() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master not found.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);
  const colIngredient = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const colCategory = getRequiredHeader_(headers, 'Category', 'Ingredients Master');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data found.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const missing = [];

  for (let i = 0; i < data.length; i++) {
    const ingredient = data[i][colIngredient - 1];
    const category = data[i][colCategory - 1];

    if (!category && ingredient) {
      missing.push(ingredient);
    }
  }

  if (missing.length === 0) {
    ui.alert('All items have categories ✅');
    return;
  }

  let message = `Uncategorised Items: ${missing.length}\n\n`;

  missing.slice(0, 10).forEach(item => {
    message += `• ${item}\n`;
  });

  if (missing.length > 10) {
    message += `\n...and ${missing.length - 10} more`;
  }

  ui.alert(message);
}

/////////////////////////////////////
// CATEGORY RULES FROM SHEET
/////////////////////////////////////
function getCategoryRules_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Category Rules');

  if (!sheet) {
    throw new Error('Missing "Category Rules" sheet.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = getHeaderMap_(sheet, 1);

  const categoryCol = getRequiredHeader_(headers, 'Category', 'Category Rules');
  const matchTextCol = getRequiredHeader_(headers, 'Match Text', 'Category Rules');
  const matchTypeCol = getRequiredHeader_(headers, 'Match Type', 'Category Rules');
  const priorityCol = headers['Priority'];
  const enabledCol = headers['Enabled'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data
    .map(row => {
      const category = (row[categoryCol - 1] || '').toString().trim();
      const matchText = (row[matchTextCol - 1] || '').toString().trim().toLowerCase();
      const matchType = (row[matchTypeCol - 1] || '').toString().trim().toLowerCase();
      const priority = priorityCol ? Number(row[priorityCol - 1]) || 999 : 999;
      const enabled = enabledCol ? (row[enabledCol - 1] || '').toString().trim().toUpperCase() : 'Y';

      return {
        category,
        matchText,
        matchType,
        priority,
        enabled
      };
    })
    .filter(r => r.category && r.matchText && r.matchType && r.enabled !== 'N')
    .sort((a, b) => a.priority - b.priority);
}

/////////////////////////////////////
// EXISTING CATEGORY LOOKUP
/////////////////////////////////////
function buildExistingCategoryLookup_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');

  if (!sheet) {
    throw new Error('Ingredients Master not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      byIngredient: {},
      byCleanName: {}
    };
  }

  const headers = getHeaderMap_(sheet, 1);
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const categoryCol = getRequiredHeader_(headers, 'Category', 'Ingredients Master');
  const cleanNameCol = headers['Clean Name'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const byIngredient = {};
  const byCleanName = {};

  data.forEach(row => {
    const ingredient = (row[ingredientCol - 1] || '').toString().trim().toLowerCase();
    const category = (row[categoryCol - 1] || '').toString().trim();
    const cleanName = cleanNameCol ? (row[cleanNameCol - 1] || '').toString().trim().toLowerCase() : '';

    if (ingredient && category && !byIngredient[ingredient]) {
      byIngredient[ingredient] = category;
    }

    if (cleanName && category && !byCleanName[cleanName]) {
      byCleanName[cleanName] = category;
    }
  });

  return { byIngredient, byCleanName };
}

/////////////////////////////////////
// CATEGORY MATCH HELPER
/////////////////////////////////////
function matchesCategoryRule_(text, rule) {
  if (!text || !rule || !rule.matchText) return false;

  const source = text.toLowerCase().trim();
  const target = rule.matchText.toLowerCase().trim();

  switch (rule.matchType) {
    case 'contains':
      return source.indexOf(target) !== -1;
    case 'equals':
      return source === target;
    case 'starts_with':
      return source.startsWith(target);
    case 'ends_with':
      return source.endsWith(target);
    default:
      return false;
  }
}

/////////////////////////////////////
// AUTO-SUGGEST PRODUCT CATEGORIES
/////////////////////////////////////
function autoSuggestIngredientCategories() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Ingredients Master sheet not found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('No ingredient rows found.');
    return;
  }

  const headers = getHeaderMap_(sheet, 1);
  const ingredientCol = getRequiredHeader_(headers, 'Ingredient', 'Ingredients Master');
  const categoryCol = getRequiredHeader_(headers, 'Category', 'Ingredients Master');
  const cleanNameCol = headers['Clean Name'];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const existingLookup = buildExistingCategoryLookup_();
  const rules = getCategoryRules_();

  let updated = 0;
  const output = [];

  for (let i = 0; i < data.length; i++) {
    const ingredient = (data[i][ingredientCol - 1] || '').toString().trim();
    const cleanName = cleanNameCol ? (data[i][cleanNameCol - 1] || '').toString().trim() : '';
    const existingCategory = (data[i][categoryCol - 1] || '').toString().trim();

    if (existingCategory) {
      output.push([existingCategory]);
      continue;
    }

    let suggestedCategory = '';

    const ingredientKey = ingredient.toLowerCase();
    const cleanNameKey = cleanName.toLowerCase();
    const combinedText = `${ingredient} ${cleanName}`.trim().toLowerCase();

    if (ingredientKey && existingLookup.byIngredient[ingredientKey]) {
      suggestedCategory = existingLookup.byIngredient[ingredientKey];
    } else if (cleanNameKey && existingLookup.byCleanName[cleanNameKey]) {
      suggestedCategory = existingLookup.byCleanName[cleanNameKey];
    } else {
      for (const rule of rules) {
        if (matchesCategoryRule_(combinedText, rule)) {
          suggestedCategory = rule.category;
          break;
        }
      }
    }

    if (suggestedCategory) updated++;
    output.push([suggestedCategory]);
  }

  sheet.getRange(2, categoryCol, output.length, 1).setValues(output);

  ui.alert(`Category auto-suggest complete.\n\nUpdated blank category cells: ${updated}`);
}