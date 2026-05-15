/////////////////////////////////////
// TEST PACK SIZE COMPATIBILITY WRAPPER
/////////////////////////////////////

function testPackSizeCompatibilityWrapper() {
  const parsed = parsePackSizeToUnitsStandard_('2-2.27ltr');

  SpreadsheetApp.getUi().alert(
    'Display Pack Size: ' +
      parsed.displayPackSize +
      '\n' +
      'Pack Qty: ' +
      parsed.packQty +
      '\n' +
      'Base Unit: ' +
      parsed.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      parsed.unitPerCase +
      '\n' +
      'Review Flag: ' +
      parsed.reviewFlag,
  );
}

/////////////////////////////////////
// TEST PILGRIM SPLIT PACK FORMAT
/////////////////////////////////////

function testPilgrimSplitPackSizeStandard() {
  const parsed = parsePackSizeStandard_({
    caseSize: '4',
    packSize: '2.5kg',
  });

  SpreadsheetApp.getUi().alert(
    'Display Pack Size: ' +
      parsed.displayPackSize +
      '\n' +
      'Pack Qty: ' +
      parsed.packQty +
      '\n' +
      'Base Unit: ' +
      parsed.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      parsed.unitPerPackCase +
      '\n' +
      'Review Flag: ' +
      parsed.reviewFlag,
  );
}

/////////////////////////////////////
// TEST PILGRIM STANDARD PACK WRAPPER
/////////////////////////////////////

function testPilgrimStandardPackWrapper() {
  const parsed = buildPilgrimStandardPackSize_('6', '1ltr');

  SpreadsheetApp.getUi().alert(
    'Pack Size: ' +
      parsed.pack_size +
      '\n' +
      'Pack Qty: ' +
      parsed.packQty +
      '\n' +
      'Base Unit: ' +
      parsed.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      parsed.unitPerPackCase +
      '\n' +
      'Review Flag: ' +
      parsed.reviewFlag,
  );
}

/////////////////////////////////////
// TEST SUPPLIER PACK WRAPPERS
/////////////////////////////////////

function testSupplierPackWrappers() {
  const bidfood = buildBidfoodStandardPackSize_('2-2.27ltr');
  const pilgrim = buildPilgrimStandardPackSize_('4', '2.5kg');

  SpreadsheetApp.getUi().alert(
    'BIDFOOD\n' +
      'Pack Size: ' +
      bidfood.pack_size +
      '\n' +
      'Pack Qty: ' +
      bidfood.packQty +
      '\n' +
      'Base Unit: ' +
      bidfood.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      bidfood.unitPerPackCase +
      '\n' +
      'Review Flag: ' +
      bidfood.reviewFlag +
      '\n\n' +
      'PILGRIM\n' +
      'Pack Size: ' +
      pilgrim.pack_size +
      '\n' +
      'Pack Qty: ' +
      pilgrim.packQty +
      '\n' +
      'Base Unit: ' +
      pilgrim.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      pilgrim.unitPerPackCase +
      '\n' +
      'Review Flag: ' +
      pilgrim.reviewFlag,
  );
}

/////////////////////////////////////
// TEST LIVE PACK PARSER USED BY REVIEW
/////////////////////////////////////

function testLivePackParserUsedByReview() {
  const tests = [
    '2-2.27ltr',
    '2-2.27l',
    '2-2.271tr',
    '2-5l',
    '2-5ltr',
    '10x400m1',
    '11t',
    'Dozen',
    '25-170-200',
  ];

  const lines = tests.map(function (packSize) {
    const parsed = parsePackSizeToUnits_(packSize);

    return (
      packSize +
      '\n' +
      'Display: ' +
      parsed.displayPackSize +
      '\n' +
      'Pack Qty: ' +
      parsed.packQty +
      '\n' +
      'Base Unit: ' +
      parsed.baseUnit +
      '\n' +
      'Unit Per Case: ' +
      parsed.unitPerCase +
      '\n' +
      'Review Flag: ' +
      parsed.reviewFlag +
      '\n' +
      'Notes: ' +
      parsed.notes
    );
  });

  SpreadsheetApp.getUi().alert(lines.join('\n\n'));
}

/////////////////////////////////////
// TEST PACK SIZE CLEANER
/////////////////////////////////////

function testCleanPackSizeStandardText() {
  SpreadsheetApp.getUi().alert(
    '2-2.27l → ' +
      cleanPackSizeStandardText_('2-2.27l') +
      '\n' +
      '2-5l → ' +
      cleanPackSizeStandardText_('2-5l') +
      '\n' +
      '24x330ml → ' +
      cleanPackSizeStandardText_('24x330ml') +
      '\n' +
      '1-6roll → ' +
      cleanPackSizeStandardText_('1-6roll'),
  );
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
    'Notes',
  ];

  const tests = getPackSizeStandardTests_();
  const output = [headers];

  tests.forEach(function (test) {
    const parsed = parsePackSizeStandard_({
      packSize: test.packSize,
      caseSize: test.caseSize,
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
      parsed.notes,
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
      packSize: 'Dozen',
      caseSize: '',
      expectedDisplayPackSize: 'dozen',
      expectedPackQty: 12,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 12,
    },
    {
      packSize: '15 Dozen',
      caseSize: '',
      expectedDisplayPackSize: '15dozen',
      expectedPackQty: 180,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 180,
    },
    {
      packSize: '2-2.27ltr',
      caseSize: '',
      expectedDisplayPackSize: '2x2.27ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4540,
    },
    {
      packSize: '2-2.27l',
      caseSize: '',
      expectedDisplayPackSize: '2x2.27ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4540,
    },
    {
      packSize: '2-2.271tr',
      caseSize: '',
      expectedDisplayPackSize: '2x2.27ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4540,
    },
    {
      packSize: '2-5l',
      caseSize: '',
      expectedDisplayPackSize: '2x5ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 10000,
    },
    {
      packSize: '1-51tr',
      caseSize: '',
      expectedDisplayPackSize: '1x5ltr',
      expectedPackQty: 1,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 5000,
    },
    {
      packSize: '2-51tr',
      caseSize: '',
      expectedDisplayPackSize: '2x5ltr',
      expectedPackQty: 2,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 10000,
    },
    {
      packSize: '10x400m1',
      caseSize: '',
      expectedDisplayPackSize: '10x400ml',
      expectedPackQty: 10,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 4000,
    },
    {
      packSize: '11t',
      caseSize: '',
      expectedDisplayPackSize: '1ltr',
      expectedPackQty: 1,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 1000,
    },
    {
      packSize: '25-170-200',
      caseSize: '',
      expectedDisplayPackSize: '25x170-200g',
      expectedPackQty: 25,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 25,
    },
    {
      packSize: '4x2.5kg',
      caseSize: '',
      expectedDisplayPackSize: '4x2.5kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10000,
    },
    {
      packSize: '2.5kg',
      caseSize: '4',
      expectedDisplayPackSize: '4x2.5kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10000,
    },
    {
      packSize: '2.16kg',
      caseSize: '5',
      expectedDisplayPackSize: '5x2.16kg',
      expectedPackQty: 5,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 10800,
    },
    {
      packSize: '24x330ml',
      caseSize: '',
      expectedDisplayPackSize: '24x330ml',
      expectedPackQty: 24,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 7920,
    },
    {
      packSize: '24x250ml',
      caseSize: '',
      expectedDisplayPackSize: '24x250ml',
      expectedPackQty: 24,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000,
    },
    {
      packSize: '12-500ml',
      caseSize: '',
      expectedDisplayPackSize: '12x500ml',
      expectedPackQty: 12,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000,
    },
    {
      packSize: '6x1ltr',
      caseSize: '',
      expectedDisplayPackSize: '6x1ltr',
      expectedPackQty: 6,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 6000,
    },
    {
      packSize: '6-2.62kg',
      caseSize: '',
      expectedDisplayPackSize: '6x2.62kg',
      expectedPackQty: 6,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 15720,
    },
    {
      packSize: '4-2.27kg',
      caseSize: '',
      expectedDisplayPackSize: '4x2.27kg',
      expectedPackQty: 4,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 9080,
    },
    {
      packSize: '24x2x28.5g',
      caseSize: '',
      expectedDisplayPackSize: '24x2x28.5g',
      expectedPackQty: 24,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1368,
    },
    {
      packSize: '5-90x20g',
      caseSize: '',
      expectedDisplayPackSize: '5x90x20g',
      expectedPackQty: 5,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 9000,
    },
    {
      packSize: '60x25g',
      caseSize: '',
      expectedDisplayPackSize: '60x25g',
      expectedPackQty: 60,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1500,
    },
    {
      packSize: '200x9g',
      caseSize: '',
      expectedDisplayPackSize: '200x9g',
      expectedPackQty: 200,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1800,
    },
    {
      packSize: '200x6g',
      caseSize: '',
      expectedDisplayPackSize: '200x6g',
      expectedPackQty: 200,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1200,
    },
    {
      packSize: '100x20g',
      caseSize: '',
      expectedDisplayPackSize: '100x20g',
      expectedPackQty: 100,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 2000,
    },
    {
      packSize: '1-6roll',
      caseSize: '',
      expectedDisplayPackSize: '6roll',
      expectedPackQty: 6,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 6,
    },
    {
      packSize: '1-2000sac',
      caseSize: '',
      expectedDisplayPackSize: '2000sac',
      expectedPackQty: 2000,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 2000,
    },
    {
      packSize: '2000',
      caseSize: '',
      expectedDisplayPackSize: '2000',
      expectedPackQty: 2000,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 2000,
    },
    {
      packSize: '48x9inch',
      caseSize: '',
      expectedDisplayPackSize: '48x9inch',
      expectedPackQty: 48,
      expectedBaseUnit: 'each',
      expectedUnitPerPackCase: 48,
    },
    {
      packSize: '1.8kg',
      caseSize: '',
      expectedDisplayPackSize: '1.8kg',
      expectedPackQty: 1,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 1800,
    },
    {
      packSize: '600g',
      caseSize: '',
      expectedDisplayPackSize: '600g',
      expectedPackQty: 1,
      expectedBaseUnit: 'g',
      expectedUnitPerPackCase: 600,
    },
    {
      packSize: '500ml',
      caseSize: '',
      expectedDisplayPackSize: '500ml',
      expectedPackQty: 1,
      expectedBaseUnit: 'ml',
      expectedUnitPerPackCase: 500,
    },
  ];
}
