/////////////////////////////////////
// BUILD PARSED ROWS FROM EXTRACTED LINES
/////////////////////////////////////

function buildParsedRowsFromExtractedLines_(fileId) {
  const ss = SpreadsheetApp.getActive();
  const extractedSheet = ss.getSheetByName('PDF Extracted Lines');
  const parsedSheet = ss.getSheetByName('PDF Parsed Rows');

  if (!extractedSheet) throw new Error('Sheet "PDF Extracted Lines" not found.');
  if (!parsedSheet) throw new Error('Sheet "PDF Parsed Rows" not found.');
  if (!fileId) throw new Error('Missing fileId.');

  const extractedHeaders = getHeaderMap_(extractedSheet, 1);
  const parsedHeaders = getHeaderMap_(parsedSheet, 1);

  /////////////////////////////////////
  // REQUIRED HEADERS - NEW STANDARD
  /////////////////////////////////////

  [
    'Upload Time',
    'File Name',
    'Supplier',
    'Site',
    'Drive File ID',
    'Row No',
    'Source Start Line',
    'Source End Line',
    'Cases',
    'Units / Weight',
    'Description',
    'Pack Size',
    'Item Code',
    'Unit Price',
    'Line Total',
    'VAT',
    'VAT Total',
    'Raw Line',
    'Status',
    'Notes'
  ].forEach(function(headerName) {
    getRequiredHeader_(parsedHeaders, headerName, 'PDF Parsed Rows');
  });

  /////////////////////////////////////
  // GET EXTRACTED LINES
  /////////////////////////////////////

  const extractedRows = getExtractedLinesForFile_(extractedSheet, extractedHeaders, fileId);

  clearParsedRowsForFile_(parsedSheet, fileId);

  if (!extractedRows.length) {
    return {
      fileId: fileId,
      sourceRows: 0,
      rowsWritten: 0
    };
  }

  /////////////////////////////////////
  // BUILD PARSED ITEMS
  /////////////////////////////////////

  const supplierName = (extractedRows[0].supplier || '')
    .toString()
    .trim()
    .toLowerCase();

  let parsedItems = [];

  if (supplierName === 'bidfood') {
    parsedItems = buildBidfoodParsedRowsFromRawLines_(extractedRows);
  } else {
    parsedItems = extractedRows.map(function(item, index) {
      return {
        'Upload Time': item.uploadTime,
        'File Name': item.fileName,
        'Supplier': item.supplier,
        'Site': item.site,
        'Drive File ID': item.fileId,

        'Row No': index + 1,
        'Source Start Line': item.lineNo,
        'Source End Line': item.lineNo,

        'Cases': '',
        'Units / Weight': '',
        'Description': item.rawLine,
        'Pack Size': '',
        'Item Code': '',
        'Unit Price': '',
        'Line Total': '',
        'VAT': '',
        'VAT Total': '',

        'Raw Line': item.rawLine,
        'Status': 'RAW',
        'Notes': 'Fallback raw line'
      };
    });
  }

  if (!parsedItems.length) {
    SpreadsheetApp.getUi().alert(
      'No parsed rows were created for this file.\n\nCheck PDF Extracted Lines and supplier raw line parsing.'
    );

    return {
      fileId: fileId,
      sourceRows: extractedRows.length,
      rowsWritten: 0,
      source: 'empty-output'
    };
  }

  /////////////////////////////////////
  // WRITE OUTPUT BY HEADERS
  /////////////////////////////////////

  const width = parsedSheet.getLastColumn();
  const startRow = parsedSheet.getLastRow() + 1;

  const output = parsedItems.map(function(item, index) {
    const row = new Array(width).fill('');

    Object.keys(parsedHeaders).forEach(function(headerName) {
      if (Object.prototype.hasOwnProperty.call(item, headerName)) {
        row[parsedHeaders[headerName] - 1] = item[headerName];
      }
    });

    // Safety fallback for core meta
    row[parsedHeaders['Upload Time'] - 1] = row[parsedHeaders['Upload Time'] - 1] || extractedRows[0].uploadTime;
    row[parsedHeaders['File Name'] - 1] = row[parsedHeaders['File Name'] - 1] || extractedRows[0].fileName;
    row[parsedHeaders['Supplier'] - 1] = row[parsedHeaders['Supplier'] - 1] || extractedRows[0].supplier;
    row[parsedHeaders['Site'] - 1] = row[parsedHeaders['Site'] - 1] || extractedRows[0].site;
    row[parsedHeaders['Drive File ID'] - 1] = row[parsedHeaders['Drive File ID'] - 1] || fileId;
    row[parsedHeaders['Row No'] - 1] = row[parsedHeaders['Row No'] - 1] || index + 1;

    return row;
  });

  parsedSheet.getRange(startRow, 1, output.length, width).setValues(output);

  return {
    fileId: fileId,
    sourceRows: extractedRows.length,
    rowsWritten: output.length
  };
}


/////////////////////////////////////
// GET EXTRACTED LINES FOR FILE
/////////////////////////////////////

function getExtractedLinesForFile_(sheet, headerMap, fileId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const uploadTimeCol = getRequiredHeader_(headerMap, 'Upload Time', 'PDF Extracted Lines');
  const fileNameCol = getRequiredHeader_(headerMap, 'File Name', 'PDF Extracted Lines');
  const supplierCol = getRequiredHeader_(headerMap, 'Supplier', 'PDF Extracted Lines');
  const siteCol = getRequiredHeader_(headerMap, 'Site', 'PDF Extracted Lines');
  const fileIdCol = getRequiredHeader_(headerMap, 'Drive File ID', 'PDF Extracted Lines');
  const lineNoCol = getRequiredHeader_(headerMap, 'Line No', 'PDF Extracted Lines');
  const rawLineCol = getRequiredHeader_(headerMap, 'Raw Line', 'PDF Extracted Lines');

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  return data
    .filter(function(row) {
      return (row[fileIdCol - 1] || '').toString().trim() === fileId;
    })
    .map(function(row) {
      return {
        uploadTime: row[uploadTimeCol - 1],
        fileName: row[fileNameCol - 1],
        supplier: row[supplierCol - 1],
        site: row[siteCol - 1],
        fileId: row[fileIdCol - 1],
        lineNo: row[lineNoCol - 1],
        rawLine: row[rawLineCol - 1]
      };
    });
}


/////////////////////////////////////
// CLEAR PARSED ROWS FOR FILE
/////////////////////////////////////

function clearParsedRowsForFile_(sheet, fileId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF Parsed Rows');

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const rowsToDelete = [];

  data.forEach(function(row, i) {
    if ((row[fileIdCol - 1] || '').toString().trim() === fileId) {
      rowsToDelete.push(i + 2);
    }
  });

  rowsToDelete.reverse().forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });
}


/////////////////////////////////////
// GET PARSED ROWS FOR CONTEXT
/////////////////////////////////////

function getPdfParsedRowsForContext_(supplier, site) {
  const ss = SpreadsheetApp.getActive();
  const parsedSheet = ss.getSheetByName('PDF Parsed Rows');

  if (!parsedSheet) throw new Error('Sheet "PDF Parsed Rows" not found.');

  const headers = getHeaderMap_(parsedSheet, 1);

  const supplierCol = getRequiredHeader_(headers, 'Supplier', 'PDF Parsed Rows');
  const siteCol = getRequiredHeader_(headers, 'Site', 'PDF Parsed Rows');
  const descCol = getRequiredHeader_(headers, 'Description', 'PDF Parsed Rows');
  const casesCol = getRequiredHeader_(headers, 'Cases', 'PDF Parsed Rows');
  const unitCol = getRequiredHeader_(headers, 'Units / Weight', 'PDF Parsed Rows');
  const priceCol = getRequiredHeader_(headers, 'Unit Price', 'PDF Parsed Rows');

  const lastRow = parsedSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = parsedSheet.getRange(2, 1, lastRow - 1, parsedSheet.getLastColumn()).getValues();

  const supplierNeedle = (supplier || '').toString().trim().toLowerCase();
  const siteNeedle = (site || '').toString().trim().toLowerCase();

  return data
    .filter(function(row) {
      const rowSupplier = (row[supplierCol - 1] || '').toString().trim().toLowerCase();
      const rowSite = (row[siteCol - 1] || '').toString().trim().toLowerCase();

      if (supplierNeedle && rowSupplier !== supplierNeedle) return false;
      if (siteNeedle && rowSite !== siteNeedle) return false;

      return true;
    })
    .map(function(row) {
      return {
        description: row[descCol - 1] || '',
        cases: row[casesCol - 1] || '',
        unitsWeight: row[unitCol - 1] || '',
        unitPrice: row[priceCol - 1] || ''
      };
    });
}