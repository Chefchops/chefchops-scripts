

/////////////////////////////////////
// VALUE / UNIT HELPERS
/////////////////////////////////////

function unitToBaseUnit(unit) {
  unit = normaliseUnit(unit);
  if (unit === 'kg' || unit === 'g') return 'g';
  if (unit === 'l' || unit === 'ml') return 'ml';
  return 'each';
}

function normaliseUnit(unit) {
  unit = unit.toLowerCase();
  if (unit === 'litre' || unit === 'liter' || unit === 'ltr') return 'l';
  return unit;
}

function convertToBaseQty(qty, unit) {
  unit = normaliseUnit(unit);
  if (unit === 'kg') return qty * 1000;
  if (unit === 'g') return qty;
  if (unit === 'l') return qty * 1000;
  if (unit === 'ml') return qty;
  return qty;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  const cleaned = value.toString().replace(/[^0-9.\-]/g, '');
  const n = Number(cleaned);
  return isNaN(n) ? 0 : n;
}

function roundNumber(value, dp) {
  const factor = Math.pow(10, dp);
  return Math.round(value * factor) / factor;
}

function normaliseKeyValue(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(value);
  if (!isNaN(num)) return num.toString();
  return value.toString().trim().toLowerCase();
}

function toTitleCase(str) {
  return str
    .toLowerCase()
    .replace(/\b\w/g, c => c.toUpperCase())
    .replace(/\bKg\b/g, 'kg')
    .replace(/\bMl\b/g, 'ml')
    .replace(/\bL\b/g, 'L');
}

function cleanBidfoodIngredientName(name) {
  if (!name) return '';

  let cleaned = name.toString();

  cleaned = cleaned
    .replace(/\bSPECIAL PRICE \d+\b/ig, '')
    .replace(/\bSPECIAL PRICE\b/ig, '')
    .replace(/\bFOC\b/ig, '')
    .replace(/\bP\d+-\d+.*$/i, '');

  cleaned = cleaned
    .replace(/\b\d+\s*-\s*\d+(?:\.\d+)?\s*(kg|g|ml|l|ea|pk|sti|sac)\b/ig, '')
    .replace(/\b\d+(?:\.\d+)?\s*(kg|g|ml|l|ea|pk|sti|sac)\b/ig, '');

  cleaned = cleaned
    .replace(/\s+\d+\-\s*$/g, '')
    .replace(/-+$/g, '');

  cleaned = cleaned
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned;
}