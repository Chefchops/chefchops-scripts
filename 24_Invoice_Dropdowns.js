///////////////////////////////////
// SET SITE DROPDOWN & SUPPLIERS
///////////////////////////////////

function setInvoiceSiteAndSupplierDropdowns() {
  const ss = SpreadsheetApp.getActive();
  const invoiceSheet = ss.getSheetByName('Invoice Import');
  const suppliersSheet = ss.getSheetByName('Suppliers');
  const sitesSheet = ss.getSheetByName('Sites');
  const ui = SpreadsheetApp.getUi();

  if (!invoiceSheet || !suppliersSheet || !sitesSheet) {
    ui.alert('Missing required sheets (Invoice Import / Suppliers / Sites).');
    return;
  }

  const supplierHeaders = getHeaderMap_(suppliersSheet, 1);
  const siteHeaders = getHeaderMap_(sitesSheet, 1);

  const supplierNameCol = getRequiredHeader_(supplierHeaders, 'Supplier Name', 'Suppliers');
  const siteNameCol = getRequiredHeader_(siteHeaders, 'Site Name', 'Sites');

  const supplierLastRow = getLastRealDataRowInColumn_(suppliersSheet, supplierNameCol, 2);
  const siteLastRow = getLastRealDataRowInColumn_(sitesSheet, siteNameCol, 2);

  if (supplierLastRow < 2) {
    ui.alert('No suppliers found on Suppliers sheet.');
    return;
  }

  if (siteLastRow < 2) {
    ui.alert('No sites found on Sites sheet.');
    return;
  }

  const supplierRange = suppliersSheet.getRange(2, supplierNameCol, supplierLastRow - 1, 1);
  const siteRange = sitesSheet.getRange(2, siteNameCol, siteLastRow - 1, 1);

  const supplierRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(supplierRange, true)
    .setAllowInvalid(false)
    .build();

  const siteRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(siteRange, true)
    .setAllowInvalid(false)
    .build();

  invoiceSheet.getRange('A4').setValue('Supplier');
  invoiceSheet.getRange('B4').clearDataValidations().clearContent().setDataValidation(supplierRule);

  invoiceSheet.getRange('A5').setValue('Site');
  invoiceSheet.getRange('B5').clearDataValidations().clearContent().setDataValidation(siteRule);

  ui.alert('Site and Supplier dropdowns updated.');
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