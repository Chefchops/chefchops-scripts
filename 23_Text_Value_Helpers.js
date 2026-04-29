/////////////////////////////////////
// TEXT + VALUE HELPERS
/////////////////////////////////////

/////////////////////////////////////
// CLEAN INGREDIENT NAME
/////////////////////////////////////
function cleanIngredientName_(text) {
  let cleaned = (text || '').toString().toLowerCase().trim();

  cleaned = cleaned
    .replace(/^half case of\s+/i, '')
    .replace(/^case of\s+/i, '')
    .replace(/^of\s+/i, '');

  cleaned = cleaned
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned;
}

/////////////////////////////////////
// TO NUMBER
/////////////////////////////////////
function toNumber(value) {
  if (value === null || value === '' || typeof value === 'undefined') return NaN;
  const num = parseFloat(value.toString().replace(/[£,\s]/g, ''));
  return isNaN(num) ? NaN : num;
}

/////////////////////////////////////
// NORMALISE RAW INVOICE ROW
// Supports comma paste in col A
// or split columns A:D
/////////////////////////////////////
function normaliseRawInvoiceRow_(rawRow) {
  const firstCell = (rawRow[0] || '').toString().trim();

  // Comma mode: everything pasted into column A
  if (firstCell.indexOf(',') !== -1) {
    const parts = firstCell.split(',').map(part => (part || '').toString().trim());

    return {
      desc: parts[0] || '',
      qty: parts[1] || '',
      unit: parts[2] || '',
      price: parts[3] || ''
    };
  }

  // Split-column mode
  return {
    desc: (rawRow[0] || '').toString().trim(),
    qty: rawRow[1] || '',
    unit: (rawRow[2] || '').toString().trim(),
    price: rawRow[3] || ''
  };
}
