/////////////////////////////////////
// BIDFOOD PDF RAWLINES ADAPTER
// PDF Extracted Lines raw text -> PDF Parsed Rows-ready objects
/////////////////////////////////////

function buildBidfoodParsedRowsFromRawLines_(extractedRows) {
  if (!extractedRows || !extractedRows.length) return [];

  Logger.log('BIDFOOD EXTRACTED ROWS: ' + extractedRows.length);
  Logger.log('BIDFOOD SAMPLE ROW: ' + JSON.stringify(extractedRows[0]));

  const customRows = extractedRows
    .map(function(row, index) {
      const parsed = parseBidfoodCustomExtractorRawLine_(row, index + 1);

      if (!parsed) {
        Logger.log('BIDFOOD FAILED RAW LINE: ' + (row.rawLine || ''));
      }

      return parsed;
    })
    .filter(Boolean)
    .filter(function(row) {
      return row.description || row.packSize || row.productCode || row.unitPrice || row.lineTotal;
    });

  Logger.log('BIDFOOD PARSED ROWS CREATED: ' + customRows.length);

  return customRows;
}


/////////////////////////////////////
// PARSE BIDFOOD CUSTOM EXTRACTOR LINE
/////////////////////////////////////

function parseBidfoodCustomExtractorRawLine_(row, lineNo) {
  const rawLine = (row.rawLine || '').toString().trim();
  if (!rawLine) return null;

  Logger.log('BIDFOOD RAW LINE: ' + rawLine);

  const parts = rawLine.split('|').map(function(p) {
    return p.trim();
  });

  Logger.log('BIDFOOD PARTS ' + parts.length + ': ' + JSON.stringify(parts));

  let caseQty = '';
  let unitsWeight = '';
  let description = '';
  let packSize = '';
  let itemCode = '';
  let unitPrice = '';
  let lineTotal = '';

  /////////////////////////////////////
  // FORMAT A:
  // caseQty | unitsWeight | description | packSize | itemCode | unitPrice | lineTotal
  /////////////////////////////////////

  if (parts.length >= 7) {
    caseQty = parts[0] || '';
    unitsWeight = parts[1] || '';
    description = parts[2] || '';
    packSize = parts[3] || '';
    itemCode = parts[4] || '';
    unitPrice = parts[5] || '';
    lineTotal = parts[6] || '';
  }

  /////////////////////////////////////
  // FORMAT B:
  // qtyOrUnits | description | packSize | itemCode | unitPrice | lineTotal
  /////////////////////////////////////

  else if (parts.length === 6) {
    const first = parts[0] || '';

    if (/units?|weight/i.test(first)) {
      unitsWeight = first;
    } else {
      caseQty = first;
    }

    description = parts[1] || '';
    packSize = parts[2] || '';
    itemCode = parts[3] || '';
    unitPrice = parts[4] || '';
    lineTotal = parts[5] || '';
  }

  else {
    return null;
  }

  const productCode = (itemCode || '').toString().replace(/\D/g, '');

  return {
    sourceStartLine: row.sourceStartLine || '',
    sourceEndLine: row.sourceEndLine || '',

    rawLine: [
      description,
      packSize,
      unitPrice,
      caseQty,
      unitsWeight,
      productCode
    ].join(' | '),

    caseQty: caseQty,
    unitsWeight: unitsWeight,
    description: description,
    packSize: packSize,
    unitPrice: unitPrice,
    lineTotal: lineTotal,
    vat: '',
    productCode: productCode,
    lineNo: lineNo
  };
}