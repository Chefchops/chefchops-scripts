function testFirstBidfoodRowKeys() {
  const fileId = Browser.inputBox('Enter Drive File ID');

  if (!fileId || fileId === 'cancel') return;

  const json = rebuildJsonFromChunks_(fileId);
  const row = (json.bidfoodRows || [])[0];

  Logger.log(JSON.stringify(row, null, 2));

  SpreadsheetApp.getUi().alert(
    'First row keys:\n\n' + Object.keys(row).join('\n')
  );
}

function testOpenPdfFolder() {
  const folderId = '1uSoRK8QrqjRSXoBBrJWDDn2S576Js-Mv';

  try {
    const folder = DriveApp.getFolderById(folderId);
    SpreadsheetApp.getUi().alert(
      'Folder opened OK:\n\n' + folder.getName()
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Could not open folder:\n\n' + err.message
    );
  }
}

/////////////////////////////////////
// TEST COMPATIBILITY WRAPPER
/////////////////////////////////////

function testPackSizeCompatibilityWrapper() {
  const parsed = parsePackSizeToUnitsStandard_('2-2.27ltr');

  SpreadsheetApp.getUi().alert(
    'Display Pack Size: ' + parsed.displayPackSize + '\n' +
    'Pack Qty: ' + parsed.packQty + '\n' +
    'Base Unit: ' + parsed.baseUnit + '\n' +
    'Unit Per Case: ' + parsed.unitPerCase + '\n' +
    'Review Flag: ' + parsed.reviewFlag
  );
}

/////////////////////////////////////
// TEST PILGRIM SPLIT PACK FORMAT
/////////////////////////////////////

function testPilgrimSplitPackSizeStandard() {
  const parsed = parsePackSizeStandard_({
    caseSize: '4',
    packSize: '2.5kg'
  });

  SpreadsheetApp.getUi().alert(
    'Display Pack Size: ' + parsed.displayPackSize + '\n' +
    'Pack Qty: ' + parsed.packQty + '\n' +
    'Base Unit: ' + parsed.baseUnit + '\n' +
    'Unit Per Case: ' + parsed.unitPerPackCase + '\n' +
    'Review Flag: ' + parsed.reviewFlag
  );
}

/////////////////////////////////////
// PILGRIM PACK SIZE STANDARD WRAPPER
/////////////////////////////////////

function buildPilgrimStandardPackSize(caseSize, packSize) {
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
// TEST PILGRIM STANDARD PACK WRAPPER
/////////////////////////////////////

function testPilgrimStandardPackWrapper() {
  const parsed = buildPilgrimStandardPackSize_('6', '1ltr');

  SpreadsheetApp.getUi().alert(
    'Pack Size: ' + parsed.pack_size + '\n' +
    'Pack Qty: ' + parsed.packQty + '\n' +
    'Base Unit: ' + parsed.baseUnit + '\n' +
    'Unit Per Case: ' + parsed.unitPerPackCase + '\n' +
    'Review Flag: ' + parsed.reviewFlag
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
    'Pack Size: ' + bidfood.pack_size + '\n' +
    'Pack Qty: ' + bidfood.packQty + '\n' +
    'Base Unit: ' + bidfood.baseUnit + '\n' +
    'Unit Per Case: ' + bidfood.unitPerPackCase + '\n' +
    'Review Flag: ' + bidfood.reviewFlag + '\n\n' +

    'PILGRIM\n' +
    'Pack Size: ' + pilgrim.pack_size + '\n' +
    'Pack Qty: ' + pilgrim.packQty + '\n' +
    'Base Unit: ' + pilgrim.baseUnit + '\n' +
    'Unit Per Case: ' + pilgrim.unitPerPackCase + '\n' +
    'Review Flag: ' + pilgrim.reviewFlag
  );
}



