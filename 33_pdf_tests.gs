
/////////////////////////////////////
// TEST PAGE 1 LINES FROM LAST PDF
/////////////////////////////////////

function testPageOneLinesFromLastPdf() {
  const stagingSheet = SpreadsheetApp.getActive().getSheetByName('PDF Staging');
  if (!stagingSheet) throw new Error('Sheet "PDF Staging" not found.');

  const lastRow = stagingSheet.getLastRow();
  if (lastRow < 2) throw new Error('No PDF rows found.');

  const fileId = stagingSheet.getRange(lastRow, 5).getValue(); // E = Drive File ID
  if (!fileId) throw new Error('No Drive File ID found on last row.');

  const json = rebuildJsonFromChunks_(fileId);
  const doc = json.document || {};
  const page = (doc.pages || [])[0];

  if (!page) throw new Error('No pages found in document.');
  if (!page.lines || !page.lines.length) throw new Error('No lines found on page 1.');

  const sample = page.lines.slice(0, 50).map(function(line, i) {
    return {
      index: i,
      text: getLayoutTextFromDoc_(doc, line.layout)
    };
  });

  Logger.log(JSON.stringify(sample, null, 2));
}
/////////////////////////////////////
// GET LAYOUT TEXT FROM DOC
/////////////////////////////////////

function getLayoutTextFromDoc_(doc, layout) {
  if (!layout || !layout.textAnchor || !layout.textAnchor.textSegments) return '';

  return layout.textAnchor.textSegments.map(function(seg) {
    const start = Number(seg.startIndex || 0);
    const end = Number(seg.endIndex || 0);
    return (doc.text || '').substring(start, end);
  }).join('').replace(/\s+/g, ' ').trim();
}


/////////////////////////////////////
// TEST EXTRACT TABLES FROM LAST PDF
/////////////////////////////////////

function testExtractTablesFromLastPdf() {
  const stagingSheet = SpreadsheetApp.getActive().getSheetByName('PDF Staging');
  if (!stagingSheet) throw new Error('Sheet "PDF Staging" not found.');

  const lastRow = stagingSheet.getLastRow();
  if (lastRow < 2) throw new Error('No PDF rows found.');

  const fileId = stagingSheet.getRange(lastRow, 5).getValue(); // E = Drive File ID
  if (!fileId) throw new Error('No Drive File ID found on last row.');

  const json = rebuildJsonFromChunks_(fileId);
  const rows = extractTablesFromDocAI_(json.document || {});

  Logger.log('File ID: ' + fileId);
  Logger.log('Rows found: ' + rows.length);
  Logger.log(JSON.stringify(rows, null, 2));
}


function extractTablesFromDocAI_(doc) {
  const pages = doc.pages || [];
  const rows = [];

  console.log('--- TABLE EXTRACTION DEBUG ---');
  console.log('Pages:', pages.length);

  pages.forEach(function(page, pageIndex) {
    const tables = page.tables || [];
    console.log('Page', pageIndex + 1, 'tables:', tables.length);

    tables.forEach(function(table, tableIndex) {
      const headerRow = (table.headerRows && table.headerRows[0] && table.headerRows[0].cells) || [];

      console.log('Table', tableIndex + 1, 'header cells:', headerRow.length);

      headerRow.forEach(function(cell, i) {
        const text = getLayoutText_(doc, cell.layout);
        console.log('Header', i, ':', text);
      });
    });
  });

  return rows;
}
function clearPdfTestSheets_() {
  const ss = SpreadsheetApp.getActive();

  const sheetNames = [
    'PDF Staging',
    'PDF JSON Staging',
    'PDF Extracted Lines',
    'PDF Parsed Rows'
  ];

  sheetNames.forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }
  });
}

/////////////////////////////////////
// PDF TESTS
/////////////////////////////////////

function testProcessPdfRow() {
  processPdfRow(2);
}

function testReadParsedText() {
  const fileId = '1spU4qi6MQ3GU1ZIn7_LYJEKhP1s36YIV';
  const json = rebuildJsonFromChunks_(fileId);

  if (!json.document || !json.document.text) {
    throw new Error('No document.text found in rebuilt JSON.');
  }

  Logger.log(json.document.text.substring(0, 2000));
}

function testExtractTables() {
  const fileId = '1spU4qi6MQ3GU1ZIn7_LYJEKhP1s36YIV';
  const json = rebuildJsonFromChunks_(fileId);
  const tables = extractTablesFromDocAI_(json.document);

  Logger.log(JSON.stringify(tables, null, 2));
}

function testBuildExtractedLinesLatest() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('PDF JSON Staging');
  const headers = getHeaderMap_(sheet, 1);
  const fileIdCol = getRequiredHeader_(headers, 'Drive File ID', 'PDF JSON Staging');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No PDF JSON data found.');

  const fileId = sheet.getRange(lastRow, fileIdCol).getValue();

  const result = buildExtractedLinesFromPdfJson_(fileId);

  Logger.log('Latest File ID: ' + fileId);
  Logger.log(JSON.stringify(result, null, 2));
}

/////////////////////////////////////
// TEST BUILD PARSED ROWS
/////////////////////////////////////

function testBuildParsedRowsFromExtractedLines() {
  const fileId = '1spU4qi6MQ3GU1ZIn7_LYJEKhP1s36YIV';
  const result = buildParsedRowsFromExtractedLines_(fileId);
  Logger.log(JSON.stringify(result, null, 2));
}


function testDocAiStructure() {
  const fileId = '1spU4qi6MQ3GU1ZIn7_LYJEKhP1s36YIV';
  const json = rebuildJsonFromChunks_(fileId);

  Logger.log('Pages: ' + (json.document.pages || []).length);
  Logger.log('Tables: ' + JSON.stringify(json.document.pages[0].tables || []));
}


function testBuildExtractedLinesFromPdfJsonNew() {
  const fileId = '1spU4qi6MQ3GU1ZIn7_LYJEKhP1s36YIV';
  const result = buildExtractedLinesFromPdfJson_(fileId);
  Logger.log(JSON.stringify(result, null, 2));
}


function testBuildInvoiceImportFromPdf() {
  buildInvoiceImportFromPdf_();
}




