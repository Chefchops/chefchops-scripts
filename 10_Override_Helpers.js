/////////////////////////////////////
// OVERRIDE / MATCH HELPERS
/////////////////////////////////////
function buildExistingOverrideMap(rawData, existingGenerated) {
  const map = {};

  for (let i = 0; i < rawData.length; i++) {
    const rawDesc = (rawData[i][0] || '').toString().trim();
    const price = toNumber(rawData[i][1]);
    const invoiceQty = toNumber(rawData[i][2]);

    if (!rawDesc) continue;

    const itemCode = (rawData[i][3] || '').toString().trim();
    const key = makeImportRowKey(rawDesc, price, invoiceQty, itemCode);
    const row = existingGenerated[i] || [];

    map[key] = {
      overrideIngredient: (row[1] || '').toString().trim(), // E
      overridePackSize: (row[5] || '').toString().trim(),   // I
      overridePackQty: row[8] !== null && row[8] !== undefined ? row[8] : '', // L
      overrideBaseUnit: (row[11] || '').toString().trim()   // O
    };
  }

  return map;
}

function makeImportRowKey(rawDesc, price, invoiceQty, itemCode) {
  return [
    (rawDesc || '').toString().trim().toLowerCase(),
    normaliseKeyValue(price),
    normaliseKeyValue(invoiceQty),
    (itemCode || '').toString().trim()
  ].join('|||');
}

function extractProductCode(rawCode) {
  if (!rawCode) return '';

  const text = rawCode.toString();

  const longMatch = text.match(/\b\d{4,}\b/);
  if (longMatch) return longMatch[0];

  const anyMatch = text.match(/\d+/);
  return anyMatch ? anyMatch[0] : '';
}