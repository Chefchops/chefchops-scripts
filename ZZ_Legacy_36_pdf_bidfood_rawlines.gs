/////////////////////////////////////
// BIDFOOD PDF RAWLINES ADAPTER
// Format:
// Cases | Units / Weight | Description | Pack Size | Item Code | Unit Price | Line Total
/////////////////////////////////////

function buildBidfoodParsedRowsFromRawLines_(extractedRows) {
  if (!extractedRows || !extractedRows.length) return [];

  console.log('BIDFOOD EXTRACTED ROWS: ' + extractedRows.length);

  const output = [];

  extractedRows.forEach(function(row, index) {
    const parsed = parseBidfoodPdfRawLine_(row, index + 1);

    if (parsed) {
      output.push(parsed);
    } else {
      console.log('BIDFOOD SKIPPED RAW LINE: ' + JSON.stringify(row));
    }
  });

  console.log('BIDFOOD FINAL PARSED COUNT: ' + output.length);
  console.log(JSON.stringify(output.slice(0, 10), null, 2));

  return output;
}


/////////////////////////////////////
// PARSE BIDFOOD PDF RAW LINE
/////////////////////////////////////

function parseBidfoodPdfRawLine_(row, lineNo) {
  const rawLine = (row.rawLine || row['Raw Line'] || '').toString().trim();
  if (!rawLine) return null;

  const parts = rawLine.split('|').map(function(p) {
    return p.trim();
  });

  console.log('BIDFOOD PARTS ' + parts.length + ': ' + JSON.stringify(parts));

    if (parts.length !== 7) {
      console.log('BAD BIDFOOD ROW - WRONG FORMAT: ' + rawLine);
      return null;
    }

    const cases = parts[0] || '';
    const unitsWeight = parts[1] || '';
    const description = parts[2] || '';
    const packSize = parts[3] || '';
    const itemCode = parts[4] || '';
    const unitPrice = parts[5] || '';
    const lineTotal = parts[6] || '';

  
const cleanItemCode = itemCode.toString().trim();

  if (!description) {
    console.log('BAD BIDFOOD ROW - MISSING DESCRIPTION: ' + rawLine);
  }

const missingFields = [];

if (!cases && !unitsWeight) missingFields.push('Cases or Units / Weight');
if (!description) missingFields.push('Description');
if (!packSize) missingFields.push('Pack Size');
if (!cleanItemCode) missingFields.push('Item Code');
if (!unitPrice) missingFields.push('Unit Price');
if (!lineTotal) missingFields.push('Line Total');

const status = missingFields.length ? 'CHECK' : 'OK';
const notes = missingFields.length
  ? 'Missing: ' + missingFields.join(', ')
  : '';

 return {
  'Upload Time': row.uploadTime || '',
  'File Name': row.fileName || '',
  'Supplier': 'Bidfood',
  'Site': row.site || '',
  'Drive File ID': row.driveFileId || '',

  'Row No': lineNo,
  'Source Start Line': row.sourceStartLine || '',
  'Source End Line': row.sourceEndLine || '',

  'Cases': parseFloat(cases) || '',
  'Units / Weight': unitsWeight || '',
  'Description': description || '',
  'Pack Size': packSize || '',
  'Item Code': cleanItemCode ? "'" + cleanItemCode : '',
  'Unit Price': parseFloat(unitPrice) || '',
  'Line Total': parseFloat(lineTotal) || '',
  'VAT': '',
  'VAT Total': '',

  'Raw Line': rawLine,
  'Status': status,
  'Notes': notes
};
}

/////////////////////////////////////
// HIGHLIGHT PDF REVIEW MISSING FIELDS
/////////////////////////////////////

function highlightPdfReviewMissingFields() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Review');

  if (!sheet) throw new Error('Sheet "PDF Review" not found.');

  const headers = getHeaderMap_(sheet, 1);

  const columnsToCheck = [
    'Corrected Cases',
    'Corrected Units / Weight',
    'Corrected Description',
    'Corrected Pack Size',
    'Corrected Item Code',
    'Corrected Unit Price',
    'Corrected Line Total'
  ];

  const rules = sheet.getConditionalFormatRules();

  columnsToCheck.forEach(function(headerName) {
    const col = getRequiredHeader_(headers, headerName, 'PDF Review');
    const range = sheet.getRange(2, col, lastRow - 1, 1);

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISBLANK(' + columnToLetter_(col) + '2)')
        .setBackground('#f4cccc')
        .setRanges([range])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('PDF Review missing-field highlighting added.');
}

/////////////////////////////////////
// COLUMN NUMBER TO LETTER
/////////////////////////////////////

function columnToLetter_(column) {
  let temp = '';
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

/////////////////////////////////////
// PDF REVIEW PROCESS GUARDS
/////////////////////////////////////

function getUnresolvedPdfParsedRows_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('PDF Parsed Rows');

  if (!sheet) throw new Error('Sheet "PDF Parsed Rows" not found.');

  const headers = getHeaderMap_(sheet, 1);

  const statusCol = getRequiredHeader_(headers, 'Status', 'PDF Parsed Rows');
  const casesCol = getRequiredHeader_(headers, 'Cases', 'PDF Parsed Rows');
  const unitsCol = getRequiredHeader_(headers, 'Units / Weight', 'PDF Parsed Rows');
  const descCol = getRequiredHeader_(headers, 'Description', 'PDF Parsed Rows');
  const packCol = getRequiredHeader_(headers, 'Pack Size', 'PDF Parsed Rows');
  const priceCol = getRequiredHeader_(headers, 'Unit Price', 'PDF Parsed Rows');
  const notesCol = getRequiredHeader_(headers, 'Notes', 'PDF Parsed Rows');

  const fileNameCol = getRequiredHeader_(headers, 'File Name', 'PDF Parsed Rows');
  const supplierCol = getRequiredHeader_(headers, 'Supplier', 'PDF Parsed Rows');
  const rowNoCol = getRequiredHeader_(headers, 'Row No', 'PDF Parsed Rows');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const unresolved = [];

  data.forEach(function(row, index) {
    const status = (row[statusCol - 1] || '').toString().trim().toUpperCase();

    const cases = row[casesCol - 1];
    const units = row[unitsCol - 1];
    const desc = row[descCol - 1];
    const pack = row[packCol - 1];
    const price = row[priceCol - 1];
    const notes = (row[notesCol - 1] || '').toString().trim();

    const problems = [];

    if (status === 'CHECK') problems.push('Status is CHECK');
    if (!cases && !units) problems.push('Missing cases or units / weight');
    if (!desc) problems.push('Missing description');
    if (!pack) problems.push('Missing pack size');
    if (!price) problems.push('Missing unit price');

    if (problems.length) {
      unresolved.push({
        sheetRow: index + 2,
        fileName: row[fileNameCol - 1] || '',
        supplier: row[supplierCol - 1] || '',
        rowNo: row[rowNoCol - 1] || '',
        status: status,
        notes: notes,
        problems: problems
      });
    }
  });

  return unresolved;
}

/////////////////////////////////////
// REQUIRE PDF REVIEW BEFORE PROCESSING
/////////////////////////////////////

function requirePdfReviewComplete_() {
  const unresolved = getUnresolvedPdfParsedRows_();

  if (!unresolved.length) return true;

  const preview = unresolved
    .slice(0, 10)
    .map(function(item) {
      return (
        'Sheet row ' + item.sheetRow +
        ' / PDF row ' + item.rowNo +
        ' / ' + item.supplier +
        '\n- ' + item.problems.join('\n- ')
      );
    })
    .join('\n\n');

  SpreadsheetApp.getUi().alert(
    'PDF Review Required',
    'Some PDF Parsed Rows still need review before processing.\n\n' +
    'Rows needing review: ' + unresolved.length + '\n\n' +
    preview +
    (unresolved.length > 10 ? '\n\nShowing first 10 only.' : '') +
    '\n\nGo to PDF Review, fix them, then apply corrections.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  /////////////////////////////////////
  // 🔥 THIS IS THE BIT YOU ADD
  /////////////////////////////////////
  const ss = SpreadsheetApp.getActive();
  const reviewSheet = ss.getSheetByName('PDF Review');

  if (reviewSheet) {
    ss.setActiveSheet(reviewSheet);
  }

  return false;
}

/////////////////////////////////////
// TEST PDF REVIEW GATE
/////////////////////////////////////

function testPdfReviewGate() {
  requirePdfReviewComplete_();
}