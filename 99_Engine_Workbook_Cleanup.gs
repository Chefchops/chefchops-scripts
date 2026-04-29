/////////////////////////////////////
// ENGINE SYSTEM HEALTH CHECK
/////////////////////////////////////

function runEngineSystemHealthCheck() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  try {
    const checks = [
      {
        sheetName: 'Ingredients Master',
        headerRow: 1,
        trustedHeaders: ['Ingredient', 'Ingredient ID']
      },
      {
        sheetName: 'Suppliers',
        headerRow: 1,
        trustedHeaders: ['Supplier Name']
      },
      {
        sheetName: 'Sites',
        headerRow: 1,
        trustedHeaders: ['Site Name']
      },
      {
        sheetName: 'Invoice Import',
        headerRow: 7,
        trustedHeaders: [],
        specialHandler: inspectInvoiceImportHealth_
      },
      {
        sheetName: 'Price Comparison',
        headerRow: 6,
        trustedHeaders: ['Ingredient']
      }
    ];

    const reportLines = [];
    const cleanupCandidates = [];
    const protectionWarnings = [];

    checks.forEach(check => {
      const result = check.specialHandler
        ? check.specialHandler(check.sheetName)
        : inspectSheetHealthByHeader_(check.sheetName, check.trustedHeaders, check.headerRow);

      if (!result) return;

      if (result.status === 'missing_sheet') {
        reportLines.push(result.sheetName + ': sheet not found');
        return;
      }

      if (result.status === 'missing_header') {
        reportLines.push(result.sheetName + ': trusted header not found');
        return;
      }

      reportLines.push(
        result.sheetName +
        ': real row ' + result.lastRealRow +
        ', sheet row ' + result.lastSheetRow +
        ', extra rows ' + result.extraRows
      );

      if (result.extraRows > 0) {
        cleanupCandidates.push(result.sheetName);
      }

      const protectedRanges = getProtectedRangeWarnings_(result.sheetName);
      protectedRanges.forEach(w => protectionWarnings.push(w));
    });

    let message = 'Engine system health check complete.\n\n';
    message += reportLines.join('\n');

    if (cleanupCandidates.length) {
      message += '\n\nCleanup suggested for:\n- ' + cleanupCandidates.join('\n- ');
    } else {
      message += '\n\nNo trailing-row cleanup needed.';
    }

    if (protectionWarnings.length) {
      message += '\n\nProtection warnings:\n- ' + protectionWarnings.join('\n- ');
    }

    ui.alert(message);

  } catch (err) {
    ui.alert('Engine system health check failed:\n\n' + err.message);
    throw err;
  }
}

/////////////////////////////////////
// ENGINE SYSTEM CLEANUP
/////////////////////////////////////

function runEngineSystemCleanup() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  try {
    const results = [];

    results.push(cleanTrailingUnusedAreaByHeader_('Ingredients Master', 'Ingredient', 'Ingredient ID'));
    results.push(cleanTrailingUnusedAreaByHeader_('Suppliers', 'Supplier Name'));
    results.push(cleanTrailingUnusedAreaByHeader_('Sites', 'Site Name'));
    results.push(cleanInvoiceImportTrailingUnusedArea_());
    results.push(cleanTrailingUnusedAreaByHeader_('Price Comparison', 'Ingredient'));

    const lines = [];

    results.forEach(result => {
      if (!result) return;

      if (result.status === 'missing_sheet') {
        lines.push(result.sheetName + ': sheet not found');
        return;
      }

      if (result.status === 'missing_header') {
        lines.push(result.sheetName + ': no trusted header found');
        return;
      }

      if (result.status === 'nothing_to_clean') {
        lines.push(result.sheetName + ': nothing to clean');
        return;
      }

      if (result.status === 'cleaned') {
        lines.push(
          result.sheetName +
          ': cleaned ' + result.rowsCleared + ' row(s) below row ' + result.lastRealRow
        );
      }
    });

    ui.alert(
      'Engine system cleanup complete.\n\n' +
      lines.join('\n')
    );

  } catch (err) {
    ui.alert('Engine system cleanup failed:\n\n' + err.message);
    throw err;
  }
}

/////////////////////////////////////
// HEALTH CHECK + CLEANUP
/////////////////////////////////////

function runEngineHealthCheckAndCleanup() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  try {
    const checks = [
      inspectSheetHealthByHeader_('Ingredients Master', ['Ingredient', 'Ingredient ID'], 1),
      inspectSheetHealthByHeader_('Suppliers', ['Supplier Name'], 1),
      inspectSheetHealthByHeader_('Sites', ['Site Name'], 1),
      inspectInvoiceImportHealth_('Invoice Import'),
      inspectSheetHealthByHeader_('Price Comparison', ['Ingredient'], 6)
    ];

    const cleanupTargets = checks.filter(r =>
      r &&
      r.status !== 'missing_sheet' &&
      r.status !== 'missing_header' &&
      r.extraRows > 0
    );

    if (!cleanupTargets.length) {
      ui.alert('Health check complete.\n\nNo cleanup needed.');
      return;
    }

    const response = ui.alert(
      'Cleanup suggested',
      'These sheets have trailing dead rows:\n\n- ' +
      cleanupTargets.map(r => r.sheetName + ' (' + r.extraRows + ' extra)').join('\n- ') +
      '\n\nRun cleanup now?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    runEngineSystemCleanup();

  } catch (err) {
    ui.alert('Engine health check + cleanup failed:\n\n' + err.message);
    throw err;
  }
}

/////////////////////////////////////
// INSPECT SHEET HEALTH BY HEADER
/////////////////////////////////////

function inspectSheetHealthByHeader_(sheetName, trustedHeaders, headerRow) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return {
      sheetName: sheetName,
      status: 'missing_sheet'
    };
  }

  const headerMap = getHeaderMap_(sheet, headerRow || 1);

  let trustedHeader = '';
  for (let i = 0; i < trustedHeaders.length; i++) {
    if (headerMap[trustedHeaders[i]]) {
      trustedHeader = trustedHeaders[i];
      break;
    }
  }

  if (!trustedHeader) {
    return {
      sheetName: sheetName,
      status: 'missing_header'
    };
  }

  const trustedCol = headerMap[trustedHeader];
  const dataStartRow = (headerRow || 1) + 1;
  const lastSheetRow = sheet.getLastRow();
  const lastRealRow = getLastRealDataRowInColumn_(sheet, trustedCol, dataStartRow);
  const extraRows = Math.max(lastSheetRow - Math.max(lastRealRow, dataStartRow - 1), 0);

  return {
    sheetName: sheetName,
    status: 'ok',
    trustedHeader: trustedHeader,
    lastSheetRow: lastSheetRow,
    lastRealRow: lastRealRow,
    extraRows: extraRows
  };
}

/////////////////////////////////////
// INSPECT INVOICE IMPORT HEALTH
/////////////////////////////////////

function inspectInvoiceImportHealth_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return {
      sheetName: sheetName,
      status: 'missing_sheet'
    };
  }

  const startRow = 8;
  const trustedCol = 1;
  const lastSheetRow = sheet.getLastRow();
  const lastRealRow = getLastRealDataRowInColumn_(sheet, trustedCol, startRow);
  const extraRows = Math.max(lastSheetRow - Math.max(lastRealRow, startRow - 1), 0);

  return {
    sheetName: sheetName,
    status: 'ok',
    trustedHeader: 'Column A from row 8',
    lastSheetRow: lastSheetRow,
    lastRealRow: lastRealRow,
    extraRows: extraRows
  };
}

/////////////////////////////////////
// CLEAN TRAILING UNUSED AREA BY HEADER
/////////////////////////////////////

function cleanTrailingUnusedAreaByHeader_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return {
      sheetName: sheetName,
      status: 'missing_sheet'
    };
  }

  const headerNames = Array.prototype.slice.call(arguments, 1);

  if (!headerNames.length) {
    throw new Error('No trusted headers supplied for cleanup: ' + sheetName);
  }

  const headerMap = getHeaderMap_(sheet, 1);

  let trustedHeader = '';
  for (let i = 0; i < headerNames.length; i++) {
    if (headerMap[headerNames[i]]) {
      trustedHeader = headerNames[i];
      break;
    }
  }

  if (!trustedHeader) {
    return {
      sheetName: sheetName,
      status: 'missing_header'
    };
  }

  const trustedCol = headerMap[trustedHeader];
  const lastSheetRow = sheet.getLastRow();
  const lastSheetCol = sheet.getLastColumn();

  if (lastSheetRow < 2) {
    return {
      sheetName: sheetName,
      status: 'nothing_to_clean',
      lastRealRow: 1,
      rowsCleared: 0
    };
  }

  const lastRealRow = getLastRealDataRowInColumn_(sheet, trustedCol, 2);
  const firstCleanupRow = Math.max(lastRealRow + 1, 2);
  const rowsToClear = lastSheetRow - firstCleanupRow + 1;

  if (rowsToClear <= 0) {
    return {
      sheetName: sheetName,
      status: 'nothing_to_clean',
      lastRealRow: lastRealRow,
      rowsCleared: 0
    };
  }

  const range = sheet.getRange(firstCleanupRow, 1, rowsToClear, lastSheetCol);
  range.clearContent();
  range.clearDataValidations();
  range.removeCheckboxes();

  return {
    sheetName: sheetName,
    status: 'cleaned',
    lastRealRow: lastRealRow,
    rowsCleared: rowsToClear
  };
}

/////////////////////////////////////
// CLEAN INVOICE IMPORT TRAILING AREA
/////////////////////////////////////

function cleanInvoiceImportTrailingUnusedArea_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Import');

  if (!sheet) {
    return {
      sheetName: 'Invoice Import',
      status: 'missing_sheet'
    };
  }

  const startRow = 8;
  const trustedCol = 1;
  const lastSheetRow = sheet.getLastRow();
  const lastSheetCol = sheet.getLastColumn();

  if (lastSheetRow < startRow) {
    return {
      sheetName: 'Invoice Import',
      status: 'nothing_to_clean',
      lastRealRow: startRow - 1,
      rowsCleared: 0
    };
  }

  const lastRealRow = getLastRealDataRowInColumn_(sheet, trustedCol, startRow);
  const firstCleanupRow = Math.max(lastRealRow + 1, startRow);
  const rowsToClear = lastSheetRow - firstCleanupRow + 1;

  if (rowsToClear <= 0) {
    return {
      sheetName: 'Invoice Import',
      status: 'nothing_to_clean',
      lastRealRow: lastRealRow,
      rowsCleared: 0
    };
  }

  const range = sheet.getRange(firstCleanupRow, 1, rowsToClear, lastSheetCol);
  range.clearContent();
  range.clearDataValidations();
  range.removeCheckboxes();

  return {
    sheetName: 'Invoice Import',
    status: 'cleaned',
    lastRealRow: lastRealRow,
    rowsCleared: rowsToClear
  };
}

/////////////////////////////////////
// PROTECTED RANGE WARNINGS
/////////////////////////////////////

function getProtectedRangeWarnings_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  const warnings = [];

  if (!sheet) return warnings;

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  protections.forEach(p => {
    const range = p.getRange();
    if (!range) return;

    const a1 = range.getA1Notation();
    const fullColumnLike =
      /^[A-Z]+:[A-Z]+$/.test(a1) ||
      /^[A-Z]+[0-9]*:[A-Z]+$/.test(a1);

    if (fullColumnLike) {
      warnings.push(sheetName + ' protected range: ' + a1);
    }
  });

  return warnings;
}

/////////////////////////////////////
// GET LAST REAL DATA ROW IN COLUMN
/////////////////////////////////////

function getLastRealDataRowInColumn_(sheet, columnNumber, startRow) {
  const firstRow = startRow || 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) return firstRow - 1;

  const values = sheet.getRange(firstRow, columnNumber, lastRow - firstRow + 1, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if ((values[i][0] || '').toString().trim() !== '') {
      return firstRow + i;
    }
  }

  return firstRow - 1;
}