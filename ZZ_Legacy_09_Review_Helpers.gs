/////////////////////////////////////
// REVIEW HELPERS
/////////////////////////////////////
function getEnhancedReviewFlag(parsed, rawDesc, price, invoiceQty) {
  if (!rawDesc) return 'CHECK';

  if (!parsed.suggestedIngredient || !parsed.packQty || !parsed.baseUnit) {
    return parsed.reviewFlag === 'CHECK PACK' ? 'CHECK PACK' : 'CHECK';
  }

  if (price > 0 && invoiceQty > 0 && parsed.packSizeDisplay === 'per kg' && parsed.baseUnit !== 'g') {
    return 'CHECK WEIGHT';
  }

  if (/\bFOC\b/i.test(rawDesc) || /\bFREE\b/i.test(rawDesc)) {
    return 'CHECK FOC';
  }

  if (price === 0 && !/\bFOC\b/i.test(rawDesc)) {
    return 'CHECK PRICE';
  }

  return parsed.reviewFlag || 'OK';
}

function getFinalReviewFlag(finalIngredient, finalPackSize, finalPackQty, finalBaseUnit, currentReviewFlag) {
  if (!finalIngredient || !finalPackSize || finalPackQty === '' || !finalBaseUnit) {
    return currentReviewFlag && currentReviewFlag !== 'OK' ? currentReviewFlag : 'CHECK';
  }
  return currentReviewFlag || 'OK';
}

function getConfirmedInvoiceContext() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Import');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Missing "Invoice Import" sheet.');
    return null;
  }

  const supplier = (sheet.getRange('B4').getValue() || '').toString().trim();
  const site = (sheet.getRange('B5').getValue() || '').toString().trim();

  if (!supplier) {
    ui.alert('Please select a Supplier in B4.');
    return null;
  }

  if (!site) {
    ui.alert('Please select a Site in B5.');
    return null;
  }

  return { supplier, site };
}

//////////////////////////////////////////
// CLEAR BUILD REVIEW HIGHLIGHTS
//////////////////////////////////////////
function clearInvoiceImportReviewHighlights_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Import');
  if (!sheet) return;

  const startRow = 8;
  const lastRow = getLastUsedRowInColumn(sheet, 4, startRow); // built block starts at D
  if (lastRow < startRow) return;

  const rowCount = lastRow - startRow + 1;

  // Clear highlight across built/output area D:AB
  sheet.getRange(startRow, 4, rowCount, 25).setBackground(null);
}

//////////////////////////////////////////
// HIGHLIGHT SPECIFIC IMPORT ROWS
//////////////////////////////////////////
function highlightInvoiceImportRows_(rowNumbers, hexColor) {
  if (!rowNumbers || rowNumbers.length === 0) return;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Import');
  if (!sheet) return;

  rowNumbers.forEach(rowNumber => {
    sheet.getRange(rowNumber, 4, 1, 25).setBackground(hexColor); // D:AB
  });
}