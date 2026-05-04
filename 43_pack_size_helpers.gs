/////////////////////////////////////
// PACK SIZE SPLITTER
// Converts supplier pack size text into:
// - packQty       = number of packs / units in the priced case
// - baseUnit      = g / ml / unit / m
// - unitPerCase   = total base units in the priced case
//
// Examples:
// 12-500ml  -> packQty 12, baseUnit ml, unitPerCase 6000
// 6x1ltr    -> packQty 6,  baseUnit ml, unitPerCase 6000
// 1ltr      -> packQty 1,  baseUnit ml, unitPerCase 1000
// 200-99    -> OCR fixed to 200x9g
/////////////////////////////////////

function parsePackSizeToUnits_(packSize) {
  const rawOriginal = (packSize || '').toString().trim();

    const raw = rawOriginal
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/×/g, 'x')
    .replace(/–/g, '-')
    .replace(/—/g, '-')

    // OCR fix: Pilgrim sometimes reads 1ltr as 11tr, 2ltr as 21tr, etc.
    // Must run before normal ltr -> l conversion.
    .replace(/^(\d+)1tr$/g, '$1ltr')

    .replace(/ltr/g, 'l')
    .replace(/lt/g, 'l')
    .replace(/litre/g, 'l')
    .replace(/litres/g, 'l')
    .replace(/liter/g, 'l')
    .replace(/liters/g, 'l')
    .trim();

  const result = {
    packQty: '',
    baseUnit: '',
    unitPerCase: '',
    reviewFlag: 'OK',
    notes: ''
  };

  if (!raw) {
    result.reviewFlag = 'CHECK';
    result.notes = 'Empty pack size';
    return result;
  }

  let match;

  /////////////////////////////////////
  // OCR FIXES / KNOWN BAD PACK FORMATS
  /////////////////////////////////////

  // Bidfood / Coronet sauces can OCR "200-9g" as "200-99"
  // Example: CORONET TOMATO KETCHUP pack size 200-99
  if (raw === '200-99') {
    result.packQty = 200;
    result.baseUnit = 'g';
    result.unitPerCase = 1800;
    result.reviewFlag = 'OK';
    result.notes = 'OCR corrected 200-99 to 200x9g';
    return result;
  }

  /////////////////////////////////////
  // HELPER: NORMALISE UNIT + TOTAL
  /////////////////////////////////////

  function convertQtyToBase_(qty, unit) {
    let baseUnit = unit;
    let total = Number(qty);

    if (unit === 'kg') {
      baseUnit = 'g';
      total = total * 1000;
    }

    if (unit === 'l') {
      baseUnit = 'ml';
      total = total * 1000;
    }

    if (unit === 'm') {
      baseUnit = 'm';
    }

    return {
      baseUnit: baseUnit,
      total: total
    };
  }

  function setOk_(packQty, baseUnit, unitPerCase) {
    result.packQty = packQty;
    result.baseUnit = baseUnit;
    result.unitPerCase = unitPerCase;
    result.reviewFlag = 'OK';
    result.notes = '';
    return result;
  }

  /////////////////////////////////////
  // 24x2x28.5g
  // Nested pack: outer x inner x size
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(g|kg|ml|l|m)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);
    const size = Number(match[3]);
    const unit = match[4];

    const converted = convertQtyToBase_(outer * inner * size, unit);

    return setOk_(
      outer * inner,
      converted.baseUnit,
      converted.total
    );
  }

  /////////////////////////////////////
  // 5-90x20g / 1-300x4g
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(g|kg|ml|l|m)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);
    const size = Number(match[3]);
    const unit = match[4];

    const converted = convertQtyToBase_(outer * inner * size, unit);

    return setOk_(
      outer,
      converted.baseUnit,
      converted.total
    );
  }

  /////////////////////////////////////
  // 12-500ml / 6-800g / 4-2.5kg
  // Also handles 12x500ml / 6x800g / 6x1l
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)[x-](\d+(?:\.\d+)?)(g|kg|ml|l|m)$/);

  if (match) {
    const qty = Number(match[1]);
    const size = Number(match[2]);
    const unit = match[3];

    const converted = convertQtyToBase_(qty * size, unit);

    return setOk_(
      qty,
      converted.baseUnit,
      converted.total
    );
  }

  /////////////////////////////////////
  // 1-120pk / 1-500ea / 6-100ptn
  // Dash count packs
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)(pk|ea|each|unit|units|ptn|portion|portions|sti|stick|sticks|roll|rolls|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);

    return setOk_(
      outer * inner,
      'unit',
      outer * inner
    );
  }

  /////////////////////////////////////
  // 12x24pk / 4x50ea / 10x100unit
  // Count x count packs
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(pk|ea|each|unit|units|ptn|portion|portions|sti|stick|sticks|roll|rolls|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);

    return setOk_(
      outer * inner,
      'unit',
      outer * inner
    );
  }

  /////////////////////////////////////
  // 20x125g without dash
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(g|kg|ml|l|m)$/);

  if (match) {
    const qty = Number(match[1]);
    const size = Number(match[2]);
    const unit = match[3];

    const converted = convertQtyToBase_(qty * size, unit);

    return setOk_(
      qty,
      converted.baseUnit,
      converted.total
    );
  }

  /////////////////////////////////////
  // 500ml / 1kg / 2.5kg / 1l / 1ltr
  // Single weight / volume / length
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)(g|kg|ml|l|m)$/);

  if (match) {
    const size = Number(match[1]);
    const unit = match[2];

    const converted = convertQtyToBase_(size, unit);

    return setOk_(
      1,
      converted.baseUnit,
      converted.total
    );
  }

  /////////////////////////////////////
  // 30ea / 2000pk / 36 / 2000
  // Simple count packs
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)(pk|ea|each|unit|units|ptn|portion|portions|sti|stick|sticks|roll|rolls|can|cans|btl|btls|sac|sachet|sachets|box|boxes|s)?$/);

  if (match) {
    const qty = Number(match[1]);

    if (qty > 0) {
      return setOk_(
        qty,
        'unit',
        qty
      );
    }
  }

  /////////////////////////////////////
  // 15dozen / 15doz / 15dz
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)(dozen|doz|dz)$/);

  if (match) {
    const qty = Number(match[1]) * 12;

    return setOk_(
      qty,
      'unit',
      qty
    );
  }

  /////////////////////////////////////
  // perkg / per kilo
  /////////////////////////////////////

  if (
    raw === 'perkg' ||
    raw === 'kg' ||
    raw === '1kgperkg'
  ) {
    return setOk_(
      1,
      'g',
      1000
    );
  }

  /////////////////////////////////////
  // each / single / unit
  /////////////////////////////////////

  if (
    raw === 'each' ||
    raw === 'ea' ||
    raw === 'single' ||
    raw === 'unit' ||
    raw === '1each' ||
    raw === '1ea' ||
    raw === '1unit'
  ) {
    return setOk_(
      1,
      'unit',
      1
    );
  }

  /////////////////////////////////////
  // FALLBACK
  /////////////////////////////////////

  result.reviewFlag = 'CHECK PACK SIZE';
  result.notes = 'Unrecognised pack size format: ' + rawOriginal;
  return result;
}