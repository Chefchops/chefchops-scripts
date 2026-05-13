/////////////////////////////////////
// PACK SIZE STANDARD
/////////////////////////////////////

function parsePackSizeStandard_(input) {
  const source = normalisePackSizeInput_(input);

  const result = {
    rawPackSize: source.rawPackSize,
    rawCaseSize: source.rawCaseSize,
    cleanedPackSize: '',
    displayPackSize: '',
    packQty: '',
    baseUnit: '',
    unitPerPackCase: '',
    reviewFlag: 'OK',
    notes: ''
  };

  let raw = cleanPackSizeStandardText_(source.rawPackSize);
  let caseSize = cleanPackSizeStandardText_(source.rawCaseSize);

  if (!raw && caseSize) {
    raw = caseSize;
    caseSize = '';
  }

  if (!raw) {
    result.reviewFlag = 'CHECK PACK SIZE';
    result.notes = 'Empty pack size';
    return result;
  }

  if (caseSize) {
    raw = caseSize + 'x' + raw;
  }

  raw = applyPackSizeOcrFixes_(raw);

  result.cleanedPackSize = raw;

  const parsed = parseCleanPackSizeStandard_(raw);

  if (!parsed.ok) {
    result.reviewFlag = 'CHECK PACK SIZE';
    result.notes = parsed.notes || 'Unrecognised pack size format: ' + result.rawPackSize;
    return result;
  }

  result.displayPackSize = parsed.displayPackSize;
  result.packQty = parsed.packQty;
  result.baseUnit = parsed.baseUnit;
  result.unitPerPackCase = parsed.unitPerPackCase;
  result.notes = parsed.notes || '';

  return result;
}

/////////////////////////////////////
// NORMALISE INPUT
/////////////////////////////////////

function normalisePackSizeInput_(input) {
  if (input && typeof input === 'object') {
    return {
      rawPackSize: input.packSize || input.pack_size || '',
      rawCaseSize: input.caseSize || input.case_size || ''
    };
  }

  return {
    rawPackSize: input || '',
    rawCaseSize: ''
  };
}

/////////////////////////////////////
// CLEAN PACK SIZE TEXT
/////////////////////////////////////

function cleanPackSizeStandardText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/\u00d7/g, 'x')
    .replace(/\u2013/g, '-')
    .replace(/\u2014/g, '-')
    .replace(/litres/g, 'ltr')
    .replace(/litre/g, 'ltr')
    .replace(/liters/g, 'ltr')
    .replace(/liter/g, 'ltr')
    .replace(/lt\b/g, 'ltr')
    .replace(/ltrs\b/g, 'ltr')
    .replace(/dozens/g, 'dozen')
    .replace(/bags/g, 'bag')
    .trim();
}

/////////////////////////////////////
// OCR FIXES
/////////////////////////////////////

function applyPackSizeOcrFixes_(raw) {
  const exactFixes = {
    '1-51tr': '1-5ltr',
    '2-51tr': '2-5ltr',
    '1-6rol1': '1-6roll',
    '12-50oml': '12-500ml',
    '24-50oml': '24-500ml',
    '1-2opk': '1-20pk',
    '200-99': '200-9g',
    '300-49': '300-4g',
    '100-29': '100-2g'
  };

  if (exactFixes[raw]) return exactFixes[raw];

  return raw
    .replace(/rol1/g, 'roll')
    .replace(/(\d+)o(ml|g|kg|ltr|pk|ea|ptn)$/g, '$10$2');
}

/////////////////////////////////////
// PARSE CLEAN PACK SIZE
/////////////////////////////////////

function parseCleanPackSizeStandard_(raw) {
  let match;

  match = raw.match(/^(\d+(?:\.\d+)?)dozen$/);
  if (match) {
    const qty = Number(match[1]) * 12;

    return {
      ok: true,
      displayPackSize: match[1] + 'dozen',
      packQty: qty,
      baseUnit: 'each',
      unitPerPackCase: qty,
      notes: ''
    };
  }

  match = raw.match(/^(\d+)x(\d+(?:\.\d+)?)dozen$/);
  if (match) {
    const outer = Number(match[1]);
    const dozen = Number(match[2]);
    const qty = outer * dozen * 12;

    return {
      ok: true,
      displayPackSize: outer + 'x' + dozen + 'dozen',
      packQty: outer,
      baseUnit: 'each',
      unitPerPackCase: qty,
      notes: ''
    };
  }

  match = raw.match(/^(\d+)[x-](\d+)x(\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);
  if (match) {
    const outer = Number(match[1]);
    const inner = Number(match[2]);
    const unitSize = Number(match[3]);
    const converted = convertPackUnitTotal_(inner * unitSize, match[4]);

    return {
      ok: true,
      displayPackSize: outer + 'x' + inner + 'x' + unitSize + displayUnit_(match[4]),
      packQty: outer,
      baseUnit: converted.baseUnit,
      unitPerPackCase: outer * converted.total,
      notes: ''
    };
  }

  match = raw.match(/^(\d+)[x-](\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);
  if (match) {
    const packQty = Number(match[1]);
    const unitSize = Number(match[2]);
    const converted = convertPackUnitTotal_(packQty * unitSize, match[3]);

    return {
      ok: true,
      displayPackSize: packQty + 'x' + unitSize + displayUnit_(match[3]),
      packQty: packQty,
      baseUnit: converted.baseUnit,
      unitPerPackCase: converted.total,
      notes: ''
    };
  }


/////////////////////////////////////
// 1-120pk / 1-500ea / 6-100ptn
/////////////////////////////////////

match = raw.match(/^(\d+)[x-](\d+)(pk|ea|each|unit|units|ptn|ptns|portion|portions|roll|rolls|sti|stick|sticks|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/);

if (match) {
  const outer = Number(match[1]);
  const inner = Number(match[2]);
  const qty = outer * inner;
  const packWord = normaliseEachPackWord_(match[3]);

  return {
    ok: true,
    displayPackSize: outer === 1
      ? inner + packWord
      : outer + 'x' + inner + packWord,
    packQty: qty,
    baseUnit: 'each',
    unitPerPackCase: qty,
    notes: ''
  };
}


  match = raw.match(/^(\d+)(pk|ea|each|unit|units|ptn|ptns|portion|portions|roll|rolls|sti|stick|sticks|can|cans|btl|btls|sac|sachet|sachets|box|boxes)$/);
  if (match) {
    const qty = Number(match[1]);

    return {
      ok: true,
      displayPackSize: qty + normaliseEachPackWord_(match[2]),
      packQty: qty,
      baseUnit: 'each',
      unitPerPackCase: qty,
      notes: ''
    };
  }

  match = raw.match(/^(\d+)x(\d+(?:\.\d+)?)(inch|in)$/);
  if (match) {
    const qty = Number(match[1]);

    return {
      ok: true,
      displayPackSize: qty + 'x' + match[2] + 'inch',
      packQty: qty,
      baseUnit: 'each',
      unitPerPackCase: qty,
      notes: 'Size descriptor retained: ' + match[2] + 'inch'
    };
  }

  match = raw.match(/^(\d+(?:\.\d+)?)(g|kg|ml|ltr|m)$/);
  if (match) {
    const unitSize = Number(match[1]);
    const converted = convertPackUnitTotal_(unitSize, match[2]);

    return {
      ok: true,
      displayPackSize: unitSize + displayUnit_(match[2]),
      packQty: 1,
      baseUnit: converted.baseUnit,
      unitPerPackCase: converted.total,
      notes: ''
    };
  }

  match = raw.match(/^(\d+)$/);
  if (match) {
    const qty = Number(match[1]);

    return {
      ok: true,
      displayPackSize: String(qty),
      packQty: qty,
      baseUnit: 'each',
      unitPerPackCase: qty,
      notes: ''
    };
  }

  return {
    ok: false,
    notes: 'Unrecognised pack size format: ' + raw
  };
}

/////////////////////////////////////
// UNIT CONVERSION
/////////////////////////////////////

function convertPackUnitTotal_(total, unit) {
  if (unit === 'kg') {
    return {
      baseUnit: 'g',
      total: total * 1000
    };
  }

  if (unit === 'ltr') {
    return {
      baseUnit: 'ml',
      total: total * 1000
    };
  }

  return {
    baseUnit: unit,
    total: total
  };
}

/////////////////////////////////////
// DISPLAY UNIT
/////////////////////////////////////

function displayUnit_(unit) {
  if (unit === 'ltr') return 'ltr';
  return unit;
}

/////////////////////////////////////
// NORMALISE EACH WORDS
/////////////////////////////////////

function normaliseEachPackWord_(word) {
  const value = String(word || '').toLowerCase();

  if (value === 'each') return 'ea';
  if (value === 'unit' || value === 'units') return 'ea';
  if (value === 'ptns' || value === 'portion' || value === 'portions') return 'ptn';
  if (value === 'rolls') return 'roll';
  if (value === 'sticks') return 'sti';
  if (value === 'stick') return 'sti';
  if (value === 'cans') return 'can';
  if (value === 'btls') return 'btl';
  if (value === 'sachet' || value === 'sachets') return 'sac';
  if (value === 'boxes') return 'box';

  return value;
}

/////////////////////////////////////
// TEST PACK SIZE STANDARD
/////////////////////////////////////

function testPackSizeStandard() {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'Pack Size Standard Test';

  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    'Raw Pack Size',
    'Raw Case Size',
    'Expected Display Pack Size',
    'Expected Pack Qty',
    'Expected Base Unit',
    'Expected Unit Per Pack/Case',
    'Actual Display Pack Size',
    'Actual Pack Qty',
    'Actual Base Unit',
    'Actual Unit Per Pack/Case',
    'Review Flag',
    'Pass / Fail',
    'Notes'
  ];

  const tests = getPackSizeStandardTests_();

  const output = [headers];

  tests.forEach(function(test) {
    const parsed = parsePackSizeStandard_({
      packSize: test.packSize,
      caseSize: test.caseSize
    });

    const pass =
      String(parsed.displayPackSize) === String(test.expectedDisplayPackSize) &&
      Number(parsed.packQty) === Number(test.expectedPackQty) &&
      String(parsed.baseUnit) === String(test.expectedBaseUnit) &&
      Number(parsed.unitPerPackCase) === Number(test.expectedUnitPerPackCase);

    output.push([
      test.packSize,
      test.caseSize,
      test.expectedDisplayPackSize,
      test.expectedPackQty,
      test.expectedBaseUnit,
      test.expectedUnitPerPackCase,
      parsed.displayPackSize,
      parsed.packQty,
      parsed.baseUnit,
      parsed.unitPerPackCase,
      parsed.reviewFlag,
      pass ? 'PASS' : 'FAIL',
      parsed.notes
    ]);
  });

  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  sheet.setFrozenRows(1);
sheet.getRange(1, 1, 1, output[0].length).setFontWeight('bold');

const existingFilter = sheet.getFilter();
if (existingFilter) {
  existingFilter.remove();
}

sheet.getDataRange().createFilter();
sheet.autoResizeColumns(1, output[0].length);

SpreadsheetApp.getUi().alert('Pack Size Standard Test complete.');
}

/////////////////////////////////////
// PACK SIZE STANDARD TESTS
/////////////////////////////////////

function getPackSizeStandardTests_() {
  return [
    {
      packSize: '2-2.27ltr',
      caseSize: '',
      expectedDisplayPackSize: '2x2.27ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4540
    },
    {
      packSize: '4x2.5kg',
      caseSize: '',
      expectedDisplayPackSize: '4x2.5kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10000
    },
    {
      packSize: '2.5kg',
      caseSize: '4',
      expectedDisplayPackSize: '4x2.5kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10000
    },
    {
      packSize: '24x330ml',
      caseSize: '',
      expectedDisplayPackSize: '24x330ml',
      expectedPackQty: 24,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 7920
    },
    {
      packSize: '6x1ltr',
      caseSize: '',
      expectedDisplayPackSize: '6x1ltr',
      expectedPackQty: 6,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000
    },
    {
      packSize: '15 Dozen',
      caseSize: '',
      expectedDisplayPackSize: '15dozen',
      expectedPackQty: 180,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 180
    },
    {
      packSize: '60x25g',
      caseSize: '',
      expectedDisplayPackSize: '60x25g',
      expectedPackQty: 60,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1500
    },
    {
      packSize: '24x2x28.5g',
      caseSize: '',
      expectedDisplayPackSize: '24x2x28.5g',
      expectedPackQty: 24,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1368
    },
    {
      packSize: '1-6roll',
      caseSize: '',
      expectedDisplayPackSize: '6roll',
      expectedPackQty: 6,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 6
    },
    {
      packSize: '1-2000sac',
      caseSize: '',
      expectedDisplayPackSize: '2000sac',
      expectedPackQty: 2000,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 2000
    },
    {
      packSize: '200x9g',
      caseSize: '',
      expectedDisplayPackSize: '200x9g',
      expectedPackQty: 200,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1800
    },
    {
      packSize: '2000',
      caseSize: '',
      expectedDisplayPackSize: '2000',
      expectedPackQty: 2000,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 2000
    },
    {
      packSize: '1-51tr',
      caseSize: '',
      expectedDisplayPackSize: '1x5ltr',
      expectedPackQty: 1,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 5000
    },
    {
      packSize: '24x240g',
      caseSize: '',
      expectedDisplayPackSize: '24x240g',
      expectedPackQty: 24,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 5760
    },
    {
      packSize: '48x9inch',
      caseSize: '',
      expectedDisplayPackSize: '48x9inch',
      expectedPackQty: 48,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 48
    },
        {
      packSize: '12-500ml',
      caseSize: '',
      expectedDisplayPackSize: '12x500ml',
      expectedPackQty: 12,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000
    },
    {
      packSize: '6-2.62kg',
      caseSize: '',
      expectedDisplayPackSize: '6x2.62kg',
      expectedPackQty: 6,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 15720
    },
    {
      packSize: '4-2.27kg',
      caseSize: '',
      expectedDisplayPackSize: '4x2.27kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 9080
    },
    {
      packSize: '10x400ml',
      caseSize: '',
      expectedDisplayPackSize: '10x400ml',
      expectedPackQty: 10,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4000
    },
    {
      packSize: '200x6g',
      caseSize: '',
      expectedDisplayPackSize: '200x6g',
      expectedPackQty: 200,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1200
    },
    {
      packSize: '100x20g',
      caseSize: '',
      expectedDisplayPackSize: '100x20g',
      expectedPackQty: 100,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 2000
    },
    {
      packSize: '24x2x28.5g',
      caseSize: '',
      expectedDisplayPackSize: '24x2x28.5g',
      expectedPackQty: 24,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1368
    },
    {
      packSize: '48x45g',
      caseSize: '',
      expectedDisplayPackSize: '48x45g',
      expectedPackQty: 48,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 2160
    },
    {
      packSize: '25x110g',
      caseSize: '',
      expectedDisplayPackSize: '25x110g',
      expectedPackQty: 25,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 2750
    },
    {
      packSize: '18x70g',
      caseSize: '',
      expectedDisplayPackSize: '18x70g',
      expectedPackQty: 18,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1260
    },
    {
      packSize: '16x60g',
      caseSize: '',
      expectedDisplayPackSize: '16x60g',
      expectedPackQty: 16,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 960
    },
    {
      packSize: '24x113g',
      caseSize: '',
      expectedDisplayPackSize: '24x113g',
      expectedPackQty: 24,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 2712
    },
    {
      packSize: '60x56g',
      caseSize: '',
      expectedDisplayPackSize: '60x56g',
      expectedPackQty: 60,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 3360
    },
    {
      packSize: '120x40g',
      caseSize: '',
      expectedDisplayPackSize: '120x40g',
      expectedPackQty: 120,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 4800
    },
    {
      packSize: '1.8kg',
      caseSize: '',
      expectedDisplayPackSize: '1.8kg',
      expectedPackQty: 1,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1800
    },
    {
      packSize: '600g',
      caseSize: '',
      expectedDisplayPackSize: '600g',
      expectedPackQty: 1,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 600
    },
    {
      packSize: '800g',
      caseSize: '',
      expectedDisplayPackSize: '800g',
      expectedPackQty: 1,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 800
    },
    {
      packSize: '500ml',
      caseSize: '',
      expectedDisplayPackSize: '500ml',
      expectedPackQty: 1,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 500
    },
    {
      packSize: '24x250ml',
      caseSize: '',
      expectedDisplayPackSize: '24x250ml',
      expectedPackQty: 24,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000
    },
    {
      packSize: '2.16kg',
      caseSize: '5',
      expectedDisplayPackSize: '5x2.16kg',
      expectedPackQty: 5,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10800
    }

    
  ];
}


/////////////////////////////////////
// COMPATIBILITY WRAPPER
// Matches old parsePackSizeToUnits_ output shape
/////////////////////////////////////

function parsePackSizeToUnitsStandard_(packSize) {
  const parsed = parsePackSizeStandard_(packSize);

  return {
    packQty: parsed.packQty,
    baseUnit: parsed.baseUnit,
    unitPerCase: parsed.unitPerPackCase,
    unitPerPackCase: parsed.unitPerPackCase,
    reviewFlag: parsed.reviewFlag,
    notes: parsed.notes,
    displayPackSize: parsed.displayPackSize,
    cleanedPackSize: parsed.cleanedPackSize
  };
}


/////////////////////////////////////
// PILGRIM PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildPilgrimStandardPackSize_(caseSize, packSize) {
  const parsed = parsePackSizeStandard_({
    caseSize: caseSize,
    packSize: packSize
  });

  return {
    pack_size: parsed.displayPackSize,
    packQty: parsed.packQty,
    baseUnit: parsed.baseUnit,
    unitPerPackCase: parsed.unitPerPackCase,
    reviewFlag: parsed.reviewFlag,
    notes: parsed.notes
  };
}

/////////////////////////////////////
// PILGRIM PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildPilgrimStandardPackSize_(caseSize, packSize) {
  const parsed = parsePackSizeStandard_({
    caseSize: caseSize,
    packSize: packSize
  });

  return {
    pack_size: parsed.displayPackSize,
    packQty: parsed.packQty,
    baseUnit: parsed.baseUnit,
    unitPerPackCase: parsed.unitPerPackCase,
    reviewFlag: parsed.reviewFlag,
    notes: parsed.notes
  };
}

/////////////////////////////////////
// BIDFOOD PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildBidfoodStandardPackSize_(packSize) {
  const parsed = parsePackSizeStandard_(packSize);

  return {
    pack_size: parsed.displayPackSize,
    packQty: parsed.packQty,
    baseUnit: parsed.baseUnit,
    unitPerPackCase: parsed.unitPerPackCase,
    reviewFlag: parsed.reviewFlag,
    notes: parsed.notes
  };
}