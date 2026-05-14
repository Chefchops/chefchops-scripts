/////////////////////////////////////
// PDF INVOICE HEADERS
// BIDFOOD FIRST
/////////////////////////////////////

const PDF_INVOICE_HEADERS_SHEET_NAME_ = 'PDF Invoice Headers';

/////////////////////////////////////
// SETUP PDF INVOICE HEADERS SHEET
/////////////////////////////////////

function setupPdfInvoiceHeadersSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(PDF_INVOICE_HEADERS_SHEET_NAME_);

  if (!sheet) {
    sheet = ss.insertSheet(PDF_INVOICE_HEADERS_SHEET_NAME_);
  }

  const headers = [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'Invoice Number',
    'Account Number',
    'Order Number',
    'Delivery Date',
    'Net Total',
    'VAT Total',
    'Gross Total',
    'Source',
    'Notes'
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.getUi().alert('PDF Invoice Headers sheet ready.');
}


/////////////////////////////////////
// BUILD LATEST PDF INVOICE HEADER
// FROM CLOUD RUN / DOCUMENT AI JSON
/////////////////////////////////////

function buildLatestPdfInvoiceHeader() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  let headerSheet = ss.getSheetByName(PDF_INVOICE_HEADERS_SHEET_NAME_);

  if (!headerSheet) {
    setupPdfInvoiceHeadersSheet();
    headerSheet = ss.getSheetByName(PDF_INVOICE_HEADERS_SHEET_NAME_);
  }

  const fileId = getLatestPdfExtractedDriveFileId_();

  if (!fileId) {
    ui.alert('No latest Drive File ID found.');
    return;
  }

  const json = rebuildJsonFromChunks_(fileId);

  if (!json) {
    ui.alert('No stored JSON found for latest file.');
    return;
  }

  const parsedJson = typeof json === 'string'
    ? JSON.parse(json)
    : json;
    Logger.log(JSON.stringify(parsedJson));

  const invoiceHeader = parsedJson.invoiceHeader || {};

  const headerHeaders = getHeaderMap_(headerSheet, 1);

upsertPdfInvoiceHeaderRow_(headerSheet, headerHeaders, {
  uploadTime: new Date(),
  fileName: parsedJson.fileName || '',
  supplier: parsedJson.supplier || '',
  site: parsedJson.site || '',
  siteName: invoiceHeader.siteName || '',
  driveFileId: fileId,

    invoiceNumber: invoiceHeader.invoiceNumber || '',
    accountNumber: invoiceHeader.accountNumber || '',
    orderNumber: invoiceHeader.orderNumber || '',
    deliveryDate: invoiceHeader.deliveryDate || '',
    netTotal: invoiceHeader.netTotal || '',
    vatTotal: invoiceHeader.vatTotal || '',
    grossTotal: invoiceHeader.grossTotal || '',
    source: invoiceHeader.source || 'Cloud Run invoiceHeader',
    notes: invoiceHeader.notes || 'No invoiceHeader notes'
  });

  formatPdfInvoiceHeadersSheet_();

  ui.alert(
    'Invoice header built from Cloud JSON\n\n' +
    'Supplier: ' + (parsedJson.supplier || '') + '\n' +
    'Invoice Number: ' + (invoiceHeader.invoiceNumber || 'Not found') + '\n' +
    'Delivery Date: ' + (invoiceHeader.deliveryDate || 'Not found')
  );
}


/////////////////////////////////////
// BUILD RAW TEXT FOR FILE
/////////////////////////////////////

function buildRawTextForFile_(rows, headers) {
  const possibleHeaders = [
    'Raw Line',
    'Text',
    'Line Text',
    'Extracted Text',
    'Description'
  ];

  let textParts = [];

  rows.forEach(function(row) {
    possibleHeaders.forEach(function(headerName) {
      const col = headers[headerName];

      if (col) {
        const value = row[col - 1];

        if (value !== null && value !== undefined && value !== '') {
          textParts.push(value.toString());
        }
      }
    });
  });

  return textParts.join('\n');
}


/////////////////////////////////////
// PARSE BIDFOOD HEADER
/////////////////////////////////////

function parseBidfoodInvoiceHeader_(text) {
  const clean = (text || '')
    .toString()
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s+/g, '\n')
    .trim();

  const result = {
    invoiceNumber: '',
    accountNumber: '',
    orderNumber: '',
    deliveryDate: '',
    netTotal: '',
    vatTotal: '',
    grossTotal: '',
    source: 'Bidfood header parser',
    notes: ''
  };

  result.invoiceNumber = findFirstMatch_(clean, [
    /invoice\s*(?:no|number|#)?\s*[:\-]?\s*([A-Z0-9\-]+)/i,
    /\binv(?:oice)?\s*[:\-]?\s*([A-Z0-9\-]+)/i
  ]);

  result.accountNumber = findFirstMatch_(clean, [
    /account\s*(?:no|number|ref)?\s*[:\-]?\s*([A-Z0-9\-]+)/i,
    /customer\s*(?:no|number|ref)?\s*[:\-]?\s*([A-Z0-9\-]+)/i
  ]);

  result.orderNumber = findFirstMatch_(clean, [
    /order\s*(?:no|number|#)?\s*[:\-]?\s*([A-Z0-9\-]+)/i,
    /\border\s*ref\s*[:\-]?\s*([A-Z0-9\-]+)/i
  ]);

  result.deliveryDate = findFirstMatch_(clean, [
    /delivery\s*date\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i,
    /del(?:ivery)?\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i
  ]);

  result.netTotal = findMoneyValue_(clean, [
    /net\s*total\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i,
    /\bnet\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i
  ]);

  result.vatTotal = findMoneyValue_(clean, [
    /vat\s*total\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i,
    /\bvat\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i
  ]);

  result.grossTotal = findMoneyValue_(clean, [
    /gross\s*total\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i,
    /invoice\s*total\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i,
    /total\s*due\s*[:\-]?\s*£?\s*([\d,]+\.\d{2})/i
  ]);

  const missing = [];

  if (!result.invoiceNumber) missing.push('Invoice Number');
  if (!result.deliveryDate) missing.push('Delivery Date');
  if (!result.netTotal && !result.grossTotal) missing.push('Totals');

  result.notes = missing.length
    ? 'Missing: ' + missing.join(', ')
    : 'OK';

  return result;
}


/////////////////////////////////////
// GENERIC HEADER PARSER FALLBACK
/////////////////////////////////////

function parseGenericInvoiceHeader_(text) {
  const clean = (text || '').toString();

  return {
    invoiceNumber: findFirstMatch_(clean, [
      /invoice\s*(?:no|number|#)?\s*[:\-]?\s*([A-Z0-9\-]+)/i
    ]),
    accountNumber: findFirstMatch_(clean, [
      /account\s*(?:no|number|ref)?\s*[:\-]?\s*([A-Z0-9\-]+)/i
    ]),
    orderNumber: findFirstMatch_(clean, [
      /order\s*(?:no|number|#)?\s*[:\-]?\s*([A-Z0-9\-]+)/i
    ]),
    deliveryDate: findFirstMatch_(clean, [
      /delivery\s*date\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i
    ]),
  
    netTotal: '',
    vatTotal: '',
    grossTotal: '',
    source: 'Generic header parser',
    notes: 'Generic parser used'
  };
}


/////////////////////////////////////
// UPSERT HEADER ROW
/////////////////////////////////////

function upsertPdfInvoiceHeaderRow_(sheet, headers, data) {
  const lastRow = sheet.getLastRow();
  let existingRow = 0;

  if (lastRow > 1) {
    const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', PDF_INVOICE_HEADERS_SHEET_NAME_);
    const fileIds = sheet.getRange(2, fileIdCol, lastRow - 1, 1).getValues();

    fileIds.forEach(function(row, index) {
      if ((row[0] || '').toString().trim() === data.driveFileId.toString().trim()) {
        existingRow = index + 2;
      }
    });
  }

  const rowValues = new Array(sheet.getLastColumn()).fill('');

  setRowByHeaders_(rowValues, headers, {
    'Upload Time': data.uploadTime,
    'File Name': data.fileName,
    'Supplier': data.supplier,
    'Site': data.siteName || data.site,
    'Drive File ID': data.driveFileId,
    'Invoice Number': data.invoiceNumber,
    'Account Number': data.accountNumber,
    'Order Number': data.orderNumber,
    'Delivery Date': data.deliveryDate,
    'Net Total': data.netTotal,
    'VAT Total': data.vatTotal,
    'Gross Total': data.grossTotal,
    'Source': data.source,
    'Notes': data.notes
  });

  if (existingRow) {
    sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}


/////////////////////////////////////
// HEADER PARSER HELPERS
/////////////////////////////////////

function findFirstMatch_(text, patterns) {
  for (let i = 0; i < patterns.length; i++) {
    const match = text.match(patterns[i]);
    if (match && match[1]) {
      return match[1].toString().trim();
    }
  }

  return '';
}


function findMoneyValue_(text, patterns) {
  const value = findFirstMatch_(text, patterns);

  if (!value) return '';

  const num = Number(value.toString().replace(/[£,]/g, '').trim());

  return isNaN(num) ? '' : num;
}


/////////////////////////////////////
// FORMAT HEADER SHEET
/////////////////////////////////////

function formatPdfInvoiceHeadersSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(PDF_INVOICE_HEADERS_SHEET_NAME_);

  if (!sheet) return;

  const headers = getHeaderMap_(sheet, 1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rowCount = lastRow - 1;

  if (headers['Upload Time']) {
    sheet.getRange(2, headers['Upload Time'], rowCount, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  }

  if (headers['Delivery Date']) {
    sheet.getRange(2, headers['Delivery Date'], rowCount, 1).setNumberFormat('dd/mm/yyyy');
  }

  ['Net Total', 'VAT Total', 'Gross Total'].forEach(function(header) {
    if (headers[header]) {
      sheet.getRange(2, headers[header], rowCount, 1).setNumberFormat('£0.00');
    }
  });

  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

/////////////////////////////////////
// BUILD PDF INVOICE HEADER FROM LATEST JSON
// SAFE ENTRY FOR PIPELINE FLOW
/////////////////////////////////////

function buildPdfInvoiceHeaderFromLatestJson_(fileId) {
  if (!fileId) throw new Error('Missing fileId for invoice header build.');

  const ss = SpreadsheetApp.getActive();

  const headerSheet = ss.getSheetByName('PDF Invoice Headers');
  if (!headerSheet) {
    throw new Error('Sheet "PDF Invoice Headers" not found.');
  }

  const headerHeaders = getHeaderMap_(headerSheet, 1);

  const parsedJson = getPdfJsonObjectForHeaderBuild_(fileId);
  const meta = getPdfJsonMetaForHeaderBuild_(fileId);

  const invoiceHeader =
    parsedJson.invoiceHeader ||
    parsedJson.invoice_header ||
    {};

  upsertPdfInvoiceHeaderRow_(headerSheet, headerHeaders, {
    uploadTime: new Date(),

    fileName:
      parsedJson.fileName ||
      parsedJson.file_name ||
      meta.fileName ||
      '',

    supplier:
      parsedJson.supplier ||
      meta.supplier ||
      '',

    site:
      parsedJson.site ||
      meta.site ||
      '',

    siteName:
      invoiceHeader.siteName ||
      invoiceHeader.site_name ||
      '',

    driveFileId: fileId,

    invoiceNumber:
      invoiceHeader.invoiceNumber ||
      invoiceHeader.invoice_number ||
      invoiceHeader.invoiceNo ||
      invoiceHeader.invoice_no ||
      '',

    accountNumber:
      invoiceHeader.accountNumber ||
      invoiceHeader.account_number ||
      '',

    orderNumber:
      invoiceHeader.orderNumber ||
      invoiceHeader.order_number ||
      '',

    deliveryDate:
      invoiceHeader.deliveryDate ||
      invoiceHeader.delivery_date ||
      '',

    netTotal:
      invoiceHeader.netTotal ||
      invoiceHeader.net_total ||
      '',

    vatTotal:
      invoiceHeader.vatTotal ||
      invoiceHeader.vat_total ||
      '',

    grossTotal:
      invoiceHeader.grossTotal ||
      invoiceHeader.gross_total ||
      '',

    source:
      invoiceHeader.source ||
      'Cloud Run invoiceHeader',

    notes:
      invoiceHeader.notes ||
      'Built from staged PDF JSON'
  });
}

/////////////////////////////////////
// GET PDF JSON OBJECT FOR HEADER BUILD
/////////////////////////////////////

function getPdfJsonObjectForHeaderBuild_(fileId) {
  if (!fileId) throw new Error('Missing fileId.');

  const json = rebuildJsonFromChunks_(fileId);

  if (!json) {
    throw new Error('No staged JSON found for Drive File ID: ' + fileId);
  }

  if (typeof json === 'object') {
    return json;
  }

  return JSON.parse(json);
}


/////////////////////////////////////
// GET PDF JSON META FOR HEADER BUILD
// HEADER-BASED FALLBACK
/////////////////////////////////////

function getPdfJsonMetaForHeaderBuild_(fileId) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF JSON Staging');

  if (!sheet) {
    throw new Error('Sheet "PDF JSON Staging" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  const headers = getHeaderMap_(sheet, 1);

  const driveFileIdCol = getRequiredHeader_(headers, 'Drive File ID');

  const fileNameCol = getOptionalHeader_(headers, 'File Name');
  const supplierCol = getOptionalHeader_(headers, 'Supplier');
  const siteCol = getOptionalHeader_(headers, 'Site');

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];

    const rowFileId = row[driveFileIdCol - 1]
      ? row[driveFileIdCol - 1].toString().trim()
      : '';

    if (rowFileId !== fileId) continue;

    return {
      fileName: fileNameCol ? row[fileNameCol - 1] : '',
      supplier: supplierCol ? row[supplierCol - 1] : '',
      site: siteCol ? row[siteCol - 1] : ''
    };
  }

  return {};
}