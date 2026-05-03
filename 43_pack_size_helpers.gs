/////////////////////////////////////
// PACK SIZE HELPERS
/////////////////////////////////////

function parsePackSizeToUnits_(packSize) {
  const raw = (packSize || '')
    .toString()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/×/g, 'x')
    .replace(/–/g, '-')
    .replace(/—/g, '-')
    .replace(/ltr/g, 'l')
    .replace(/lt/g, 'l')
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
  // 5-90x20g / 1-300x4g
  /////////////////////////////////////

  match = raw.match(/^(\d+)-(\d+)x(\d+(?:\.\d+)?)(g|kg|ml|l)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);
    const size = Number(match[3]);
    let unit = match[4];

    let total = outer * inner * size;

    if (unit === 'kg') {
      unit = 'g';
      total = total * 1000;
    }

    if (unit === 'l') {
      unit = 'ml';
      total = total * 1000;
    }

    result.packQty = outer;
    result.baseUnit = unit;
    result.unitPerCase = total;

    return result;
  }

  /////////////////////////////////////
  // 12-500ml / 6-800g / 4-2.5kg
  /////////////////////////////////////

  match = raw.match(/^(\d+)[x-](\d+(?:\.\d+)?)(g|kg|ml|l)$/);

  if (match) {
    const qty = Number(match[1]);
    const size = Number(match[2]);
    let unit = match[3];

    let total = qty * size;

    if (unit === 'kg') {
      unit = 'g';
      total = total * 1000;
    }

    if (unit === 'l') {
      unit = 'ml';
      total = total * 1000;
    }

    result.packQty = qty;
    result.baseUnit = unit;
    result.unitPerCase = total;

    return result;
  }

  /////////////////////////////////////
  // 1-120pk / 5-10pk / 1-500ea
  /////////////////////////////////////

  match = raw.match(/^(\d+)[x-](\d+)(pk|ea|ptn|portion|portions)$/);

  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);

    result.packQty = outer;
    result.baseUnit = 'each';
    result.unitPerCase = outer * inner;

    return result;
  }

  /////////////////////////////////////
  // 48pk / 500ea / 36ptn
  /////////////////////////////////////

  match = raw.match(/^(\d+)(pk|ea|ptn|portion|portions)$/);

  if (match) {
    const qty = Number(match[1]);

    result.packQty = qty;
    result.baseUnit = 'each';
    result.unitPerCase = qty;

    return result;
  }

  /////////////////////////////////////
  // 5kg / 500g / 750ml
  /////////////////////////////////////

  match = raw.match(/^(\d+(?:\.\d+)?)(g|kg|ml|l)$/);

  if (match) {
    const size = Number(match[1]);
    let unit = match[2];

    let total = size;

    if (unit === 'kg') {
      unit = 'g';
      total = total * 1000;
    }

    if (unit === 'l') {
      unit = 'ml';
      total = total * 1000;
    }

    result.packQty = 1;
    result.baseUnit = unit;
    result.unitPerCase = total;

    return result;
  }

  /////////////////////////////////////
// 1-1000sti / 1-6roll
// treat as each-count packs
/////////////////////////////////////

match = raw.match(/^(\d+)[x-](\d+)(sti|stick|sticks|roll|rolls)$/);

if (match) {
  const outer = Number(match[1]);
  const inner = Number(match[2]);

  result.packQty = outer;
  result.baseUnit = 'each';
  result.unitPerCase = outer * inner;

  return result;
}

  /////////////////////////////////////
  // FALLBACK
  /////////////////////////////////////

  result.reviewFlag = 'CHECK';
  result.notes = 'Unrecognised pack size format: ' + packSize;

  return result;
}