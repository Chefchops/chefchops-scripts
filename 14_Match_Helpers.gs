/////////////////////////////////////
// MATCH HELPERS
/////////////////////////////////////

/////////////////////////////////////
// NORMALISE MATCH VALUE
/////////////////////////////////////
function normaliseMatchValue_(value) {
  return (value || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/////////////////////////////////////
// NORMALISE MATCH NUMBER
/////////////////////////////////////
function normaliseMatchNumber_(value) {
  const num = Number(value);
  if (isNaN(num)) return '';
  return Number(num.toFixed(6)).toString();
}

/////////////////////////////////////
// BUILD INCOMING CODE MATCH KEY
/////////////////////////////////////
function buildIncomingCodeMatchKey_(valuesObj) {
  const itemCode = normaliseMatchValue_(valuesObj['Item Code']);
  const supplier = normaliseMatchValue_(valuesObj['Supplier']);

  if (!itemCode || !supplier) return '';
  return `CODE|||${itemCode}|||${supplier}`;
}

/////////////////////////////////////
// BUILD INCOMING FALLBACK MATCH KEY
/////////////////////////////////////
function buildIncomingFallbackMatchKey_(valuesObj) {
  const cleanName = normaliseMatchValue_(valuesObj['Clean Name']);
  const supplier = normaliseMatchValue_(valuesObj['Supplier']);
  const packQty = normaliseMatchNumber_(valuesObj['Pack Qty']);
  const baseUnit = normaliseMatchValue_(valuesObj['Base Unit']);
  const packSize = normaliseMatchValue_(valuesObj['Pack Size']);

  if (!cleanName || !supplier || !packQty || !baseUnit) return '';
  return `NAMEQTY|||${cleanName}|||${supplier}|||${packQty}|||${baseUnit}|||${packSize}`;
}

/////////////////////////////////////
// BUILD MASTER MATCH KEY
/////////////////////////////////////
function buildMasterMatchKey_(row, headerMap) {
  const itemCodeCol = getOptionalHeader_(headerMap, 'Item Code');
  const supplierCol = getRequiredHeader_(headerMap, 'Supplier', 'Ingredients Master');
  const cleanNameCol = getRequiredHeader_(headerMap, 'Clean Name', 'Ingredients Master');
  const packQtyCol = getRequiredHeader_(headerMap, 'Pack Qty', 'Ingredients Master');
  const baseUnitCol = getRequiredHeader_(headerMap, 'Base Unit', 'Ingredients Master');
  const packSizeCol = getOptionalHeader_(headerMap, 'Pack Size');

  const itemCode = itemCodeCol ? normaliseMatchValue_(row[itemCodeCol - 1]) : '';
  const supplier = normaliseMatchValue_(row[supplierCol - 1]);
  const cleanName = normaliseMatchValue_(row[cleanNameCol - 1]);
  const packQty = normaliseMatchNumber_(row[packQtyCol - 1]);
  const baseUnit = normaliseMatchValue_(row[baseUnitCol - 1]);
  const packSize = packSizeCol ? normaliseMatchValue_(row[packSizeCol - 1]) : '';

  if (itemCode && supplier) {
    return `CODE|||${itemCode}|||${supplier}`;
  }

  if (!cleanName || !supplier || !packQty || !baseUnit) return '';
  return `NAMEQTY|||${cleanName}|||${supplier}|||${packQty}|||${baseUnit}|||${packSize}`;
}

/////////////////////////////////////
// BUILD INCOMING MATCH KEY
/////////////////////////////////////
function buildIncomingMatchKey_(valuesObj) {
  const codeKey = buildIncomingCodeMatchKey_(valuesObj);
  if (codeKey) return codeKey;
  return buildIncomingFallbackMatchKey_(valuesObj);
}

/////////////////////////////////////
// FIND MATCHING INGREDIENTS MASTER ROW
/////////////////////////////////////
function findMatchingIngredientsMasterRow_(sheet, headerMap, incomingValues) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const incomingKey = buildIncomingMatchKey_(incomingValues);

  if (!incomingKey) return 0;

  for (let i = 0; i < data.length; i++) {
    const existingKey = buildMasterMatchKey_(data[i], headerMap);
    if (existingKey && existingKey === incomingKey) {
      return i + 2;
    }
  }

  return 0;
}
