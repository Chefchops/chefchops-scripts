//////////////////////////////////////////
// APPEND INVOICE ROWS TO INGREDIENTS MASTER
// HEADER-BASED ON INGREDIENTS MASTER
///////////////////////////////////////////


function appendInvoiceRowsToIngredientsMaster() {
  if (!requirePdfReviewComplete_()) return;
  const existingIds = [];
  const ss = SpreadsheetApp.getActive();

  const importSheet = ss.getSheetByName('Invoice Import');
  const masterSheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  try {
    if (!importSheet || !masterSheet) {
      ui.alert('Missing "Invoice Import" or "Ingredients Master" sheet.');
      return;
    }

    const context = getConfirmedInvoiceContext();
    if (!context) return;

    const selectedSupplier = context.supplier;
    const selectedSite = context.site;

    const startRow = 8;
    const lastRow = getLastUsedRowInColumn(importSheet, 1, startRow);

    if (lastRow < startRow) {
      ui.alert('No invoice rows found.');
      return;
    }

    const rowCount = lastRow - startRow + 1;

    /////////////////////////////////////////
    // KEEP IMPORT BLOCK POSITIONAL FOR NOW
    /////////////////////////////////////////
    const finalBlock = importSheet.getRange(startRow, 4, rowCount, 15).getValues();  // D:R
    const masterBlock = importSheet.getRange(startRow, 20, rowCount, 9).getValues(); // T:AB

    /////////////////////////////////////////
    // INGREDIENTS MASTER = HEADER-BASED
    /////////////////////////////////////////
    const MASTER_HEADER_ROW = 1;
    const masterHeaders = getHeaderMap_(masterSheet, MASTER_HEADER_ROW);

    const colIngredientId   = getRequiredHeader_(masterHeaders, 'Ingredient ID', 'Ingredients Master');
    const colIngredient     = getRequiredHeader_(masterHeaders, 'Ingredient', 'Ingredients Master');
    const colCleanName      = getRequiredHeader_(masterHeaders, 'Clean Name', 'Ingredients Master');
    const colCategory       = getOptionalHeader_(masterHeaders, 'Category');
    const colProductGroup   = getOptionalHeader_(masterHeaders, 'Product Group');
    const colSupplier       = getRequiredHeader_(masterHeaders, 'Supplier', 'Ingredients Master');
    const colPackSize       = getRequiredHeader_(masterHeaders, 'Pack Size', 'Ingredients Master');
    const colPackQty        = getRequiredHeader_(masterHeaders, 'Pack Qty', 'Ingredients Master');
    const colPackPrice      = getRequiredHeader_(masterHeaders, 'Pack Price (£)', 'Ingredients Master');
    const colBaseUnit       = getRequiredHeader_(masterHeaders, 'Base Unit', 'Ingredients Master');
    const colCostPerUnit    = getRequiredHeader_(masterHeaders, 'Cost per Unit (£)', 'Ingredients Master');
    const colItemCode       = getOptionalHeader_(masterHeaders, 'Item Code');
    const colNotes          = getOptionalHeader_(masterHeaders, 'Notes');
    const colRecipeUnitCost = getOptionalHeader_(masterHeaders, 'Recipe/Unit Cost');

    const masterLastRow = masterSheet.getLastRow();
    const masterLastCol = masterSheet.getLastColumn();
    const masterDataRows = Math.max(masterLastRow - MASTER_HEADER_ROW, 0);

    let masterData = [];
    if (masterDataRows > 0) {
      masterData = masterSheet.getRange(MASTER_HEADER_ROW + 1, 1, masterDataRows, masterLastCol).getValues();
    }

    /////////////////////////////////////////
    // COLLECT EXISTING IDS
    /////////////////////////////////////////
    masterData.forEach(row => {
      const existingId = (row[colIngredientId - 1] || '').toString().trim();
      if (existingId) existingIds.push(existingId);
    });

    /////////////////////////////////////////
    // BUILD MATCH MAPS
    /////////////////////////////////////////
    const codeMap = new Map();
    const fallbackMap = new Map();

    for (let i = 0; i < masterData.length; i++) {
      const rowNumber = i + MASTER_HEADER_ROW + 1;
      const row = masterData[i];

      const existingValues = {
        'Supplier': row[colSupplier - 1],
        'Clean Name': row[colCleanName - 1],
        'Pack Qty': row[colPackQty - 1],
        'Base Unit': row[colBaseUnit - 1],
        'Pack Size': row[colPackSize - 1],
        'Item Code': colItemCode ? row[colItemCode - 1] : ''
      };

      const codeKey = buildIncomingCodeMatchKey_(existingValues);
      const fallbackKey = buildIncomingFallbackMatchKey_(existingValues);

      if (codeKey) codeMap.set(codeKey, rowNumber);
      if (fallbackKey) fallbackMap.set(fallbackKey, rowNumber);
    }

    const rowsToAppend = [];
    const updates = [];

    let addedCount = 0;
    let updatedCount = 0;
    let skippedCount = 0;
    const priceAlertsRows = [];

    /////////////////////////////////////////
    // PROCESS IMPORT ROWS
    /////////////////////////////////////////
    for (let i = 0; i < rowCount; i++) {
      const finalRow = finalBlock[i];
      const builtRow = masterBlock[i];

      const finalIngredient = (finalRow[2] || '').toString().trim();
      const finalPackSize = (finalRow[6] || '').toString().trim();
      const finalPackQty = finalRow[9];
      const finalBaseUnit = (finalRow[12] || '').toString().trim();
      const review = (finalRow[14] || '').toString().trim();

      const cleanName = (builtRow[1] || '').toString().trim();
      const supplier = (builtRow[2] || '').toString().trim();
      const packPrice = builtRow[5];
      const costPerUnit = builtRow[7];
      const itemCode = (builtRow[8] || '').toString().trim();

      if (!finalIngredient || !cleanName || !supplier || !finalPackQty || !finalBaseUnit) {
        skippedCount++;
        continue;
      }

      if (review === 'CREDIT') {
        skippedCount++;
        continue;
      }

      if (Number(packPrice) < 0) {
        skippedCount++;
        continue;
      }

      if (supplier.toLowerCase() !== selectedSupplier.toLowerCase()) {
        ui.alert(`Supplier mismatch found.\n\nBuilt: ${supplier}\nSelected: ${selectedSupplier}`);
        return;
      }

      const incomingValues = {
        'Ingredient': finalIngredient,
        'Clean Name': cleanName,
        'Supplier': supplier,
        'Pack Size': finalPackSize,
        'Pack Qty': finalPackQty,
        'Pack Price (£)': packPrice,
        'Base Unit': finalBaseUnit,
        'Cost per Unit (£)': costPerUnit,
        'Item Code': itemCode
      };

      const codeKey = buildIncomingCodeMatchKey_(incomingValues);
      const fallbackKey = buildIncomingFallbackMatchKey_(incomingValues);

      let matchRow = null;

      if (codeKey && codeMap.has(codeKey)) matchRow = codeMap.get(codeKey);
      if (!matchRow && fallbackKey && fallbackMap.has(fallbackKey)) matchRow = fallbackMap.get(fallbackKey);

      if (matchRow) {
        const existingRow = masterSheet.getRange(matchRow, 1, 1, masterLastCol).getValues()[0];

        setRowByHeaders_(existingRow, masterHeaders, incomingValues);

        updates.push({ rowNumber: matchRow, rowValues: existingRow });
        updatedCount++;

      } else {
        const newRow = new Array(masterLastCol).fill('');
        const ingredientId = generateNextIngredientId_(existingIds);
        existingIds.push(ingredientId);

        setRowByHeaders_(newRow, masterHeaders, {
          'Ingredient ID': ingredientId,
          ...incomingValues,
          'Notes': 'Imported from Invoice Import | New Row Added'
        });

        rowsToAppend.push(newRow);
        addedCount++;
      }
    }

    /////////////////////////////////////////
    // WRITE UPDATES
    /////////////////////////////////////////
    updates.forEach(u => {
      masterSheet.getRange(u.rowNumber, 1, 1, masterLastCol).setValues([u.rowValues]);
    });

    /////////////////////////////////////////
    // APPEND NEW ROWS
    /////////////////////////////////////////
    if (rowsToAppend.length > 0) {
      const insertRow = getLastDataRowByHeader_(masterSheet, 1, 'Ingredient') + 1;
      masterSheet.getRange(insertRow, 1, rowsToAppend.length, masterLastCol).setValues(rowsToAppend);
    }

    /////////////////////////////////////////
    // FINAL MESSAGE
    /////////////////////////////////////////
    ui.alert(
      `Ingredients Master updated.\n\nAdded: ${addedCount}\nUpdated: ${updatedCount}\nSkipped: ${skippedCount}`
    );

  } catch (err) {
    ui.alert('appendInvoiceRowsToIngredientsMaster error:\n\n' + err.message);
    throw err;
  }
}

function getRawInvoiceFieldsBySupplier(rawRow, supplier) {
  const supplierKey = (supplier || '').toString().trim().toLowerCase();

  const cell = (i) => (rawRow[i] || '').toString().trim();
  const num = (i) => toNumber(rawRow[i]);

  function splitCommaRow_(text) {
    if (!text || text.indexOf(',') === -1) return null;

    const parts = text.split(',').map(v => (v || '').toString().trim());
    if (parts.length < 2) return null;

    return parts;
  }

  function looksLikeMostlyBlankSplitRow_() {
    let used = 0;
    for (let i = 0; i < Math.min(rawRow.length, 10); i++) {
      if (cell(i) !== '') used++;
    }
    return used <= 1;
  }

  function mapGenericCommaPaste_(text) {
    const parts = splitCommaRow_(text);
    if (!parts) return null;

    // safest generic assumption:
    // Description, Price, Qty, ItemCode
    // or
    // Description, Qty, Unit, Price
    if (parts.length >= 4) {
      const secondIsNumber = !isNaN(toNumber(parts[1]));
      const thirdIsNumber = !isNaN(toNumber(parts[2]));
      const fourthIsNumber = !isNaN(toNumber(parts[3]));

      // Description, Qty, Unit, Price
      if (secondIsNumber && !thirdIsNumber && fourthIsNumber) {
        return {
          itemCode: '',
          rawDesc: parts[0] || '',
          packSize: '',
          pack: '',
          size: '',
          caseSize: '',
          casesOrdered: '',
          singlesOrdered: '',
          price: toNumber(parts[3]),
          invoiceQty: toNumber(parts[1]),
          unitsWeight: '',
          qtyType: parts[2] || ''
        };
      }

      // Description, Price, Qty, ItemCode
      if (secondIsNumber && thirdIsNumber) {
        return {
          itemCode: parts[3] || '',
          rawDesc: parts[0] || '',
          packSize: '',
          pack: '',
          size: '',
          caseSize: '',
          casesOrdered: '',
          singlesOrdered: '',
          price: toNumber(parts[1]),
          invoiceQty: toNumber(parts[2]),
          unitsWeight: '',
          qtyType: ''
        };
      }
    }

    // fallback
    return {
      itemCode: '',
      rawDesc: parts[0] || '',
      packSize: '',
      pack: '',
      size: '',
      caseSize: '',
      casesOrdered: '',
      singlesOrdered: '',
      price: toNumber(parts[1]),
      invoiceQty: toNumber(parts[2]),
      unitsWeight: '',
      qtyType: parts[3] || ''
    };
  }

    /////////////////////////////////////
  // PILGRIM
  // Supports:
  // 1) split columns
  // 2) comma paste in column A
  /////////////////////////////////////
  if (supplierKey.indexOf('pilgrim') !== -1) {
    const firstCell = (rawRow[0] || '').toString().trim();

    const otherCellsUsed =
      ((rawRow[1] || '').toString().trim() !== '') ||
      ((rawRow[2] || '').toString().trim() !== '') ||
      ((rawRow[3] || '').toString().trim() !== '') ||
      ((rawRow[4] || '').toString().trim() !== '') ||
      ((rawRow[5] || '').toString().trim() !== '');

    // COMMA MODE
    if (firstCell.indexOf(',') !== -1 && !otherCellsUsed) {
      const parts = firstCell.split(',').map(p => (p || '').toString().trim());

      return {
        itemCode: parts[0] || '',
        rawDesc: parts[1] || '',
        caseSize: parts[2] || '',
        packSize: parts[3] || '',
        pack: '',
        size: '',
        casesOrdered: toNumber(parts[4]),
        singlesOrdered: toNumber(parts[5]),
        price: toNumber(parts[6]),
        invoiceQty: '',
        unitsWeight: '',
        qtyType: ''
      };
    }

    // SPLIT MODE (existing)
    return {
      itemCode: (rawRow[0] || '').toString().trim(),
      rawDesc: (rawRow[1] || '').toString().trim(),
      caseSize: (rawRow[2] || '').toString().trim(),
      packSize: (rawRow[3] || '').toString().trim(),
      pack: '',
      size: '',
      casesOrdered: toNumber(rawRow[4]),
      singlesOrdered: toNumber(rawRow[5]),
      price: toNumber(rawRow[6]),
      invoiceQty: '',
      unitsWeight: '',
      qtyType: ''
    };
  }

    /////////////////////////////////////
  // MAKRO
  // Supports:
  // 1) split columns
  //    A = Item Code
  //    B = Description
  //    C = Pack
  //    D = Size
  //    E = Price
  //    F = Qty
  //
  // 2) comma paste in column A
  //    Item Code,Description,Pack,Size,Price,Qty
  /////////////////////////////////////
  if (supplierKey.indexOf('makro') !== -1) {
    const firstCell = (rawRow[0] || '').toString().trim();

    const otherCellsUsed =
      ((rawRow[1] || '').toString().trim() !== '') ||
      ((rawRow[2] || '').toString().trim() !== '') ||
      ((rawRow[3] || '').toString().trim() !== '') ||
      ((rawRow[4] || '').toString().trim() !== '') ||
      ((rawRow[5] || '').toString().trim() !== '');

    // COMMA MODE
    if (firstCell.indexOf(',') !== -1 && !otherCellsUsed) {
      const parts = firstCell.split(',').map(p => (p || '').toString().trim());

      return {
        itemCode: parts[0] || '',
        rawDesc: parts[1] || '',
        pack: parts[2] || '',
        size: parts[3] || '',
        packSize: '',
        caseSize: '',
        casesOrdered: '',
        singlesOrdered: '',
        price: toNumber(parts[4]),
        invoiceQty: toNumber(parts[5]),
        unitsWeight: '',
        qtyType: ''
      };
    }

    // SPLIT MODE
    return {
      itemCode: cell(0),
      rawDesc: cell(1),
      pack: cell(2),
      size: cell(3),
      packSize: '',
      caseSize: '',
      casesOrdered: '',
      singlesOrdered: '',
      price: num(4),
      invoiceQty: num(5),
      unitsWeight: '',
      qtyType: ''
    };
  }

    // BIDFOOD
  // Supports both:
  // 1) split columns
  //    A = Description
  //    B = Pack Size
  //    C = Price
  //    D = Cases
  //    E = Units/Weight
  //    F = Item Code
  //
  // 2) comma paste in column A
  //    Description,Pack Size,Price,Cases,Units/Weight,Item Code
  if (supplierKey.indexOf('bidfood') !== -1) {
    const firstCell = (rawRow[0] || '').toString().trim();

    // detect comma-paste row only if rest of row is basically blank
    const otherCellsUsed =
      ((rawRow[1] || '').toString().trim() !== '') ||
      ((rawRow[2] || '').toString().trim() !== '') ||
      ((rawRow[3] || '').toString().trim() !== '') ||
      ((rawRow[4] || '').toString().trim() !== '') ||
      ((rawRow[5] || '').toString().trim() !== '');

    if (firstCell.indexOf(',') !== -1 && !otherCellsUsed) {
      const parts = firstCell.split(',').map(part => (part || '').toString().trim());

      return {
        rawDesc: parts[0] || '',
        packSize: parts[1] || '',
        price: toNumber(parts[2]),
        invoiceQty: toNumber(parts[3]),
        unitsWeight: parts[4] || '',
        qtyType: '',
        itemCode: parts[5] || ''
      };
    }

    // normal split-column mode
    return {
      rawDesc: (rawRow[0] || '').toString().trim(),
      packSize: (rawRow[1] || '').toString().trim(),
      price: toNumber(rawRow[2]),
      invoiceQty: toNumber(rawRow[3]),
      unitsWeight: (rawRow[4] || '').toString().trim(),
      qtyType: '',
      itemCode: (rawRow[5] || '').toString().trim()
    };
  }

  /////////////////////////////////////
  // FREESTONES
  // Supports:
  // 1) split columns
  //    A = Description
  //    B = Qty
  //    C = Unit
  //    D = Price
  //    E = Item Code
  //
  // 2) comma paste in column A
  //    Item Code,Description,Qty,Unit,Price
  /////////////////////////////////////
  if (supplierKey.indexOf('freeston') !== -1) {
    const firstCell = (rawRow[0] || '').toString().trim();

    const otherCellsUsed =
      ((rawRow[1] || '').toString().trim() !== '') ||
      ((rawRow[2] || '').toString().trim() !== '') ||
      ((rawRow[3] || '').toString().trim() !== '') ||
      ((rawRow[4] || '').toString().trim() !== '');

    // COMMA MODE
    if (firstCell.indexOf(',') !== -1 && !otherCellsUsed) {
      const parts = firstCell.split(',').map(p => (p || '').toString().trim());

      return {
        itemCode: parts[0] || '',
        rawDesc: parts[1] || '',
        packSize: '',
        pack: '',
        size: '',
        caseSize: '',
        casesOrdered: '',
        singlesOrdered: '',
        price: toNumber(parts[4]),
        invoiceQty: toNumber(parts[2]),
        unitsWeight: '',
        qtyType: parts[3] || ''
      };
    }

    // SPLIT MODE
    return {
      itemCode: cell(4),
      rawDesc: cell(0),
      packSize: '',
      pack: '',
      size: '',
      caseSize: '',
      casesOrdered: '',
      singlesOrdered: '',
      price: num(3),
      invoiceQty: num(1),
      unitsWeight: '',
      qtyType: cell(2)
    };
  }
  /////////////////////////////////////
  // HAZELS
  // Supports:
  // 1) split columns
  //    A = Ref / Item Code
  //    B = Description
  //    C = Packs (ignore)
  //    D = Qty
  //    E = Price Per
  //    F = Unit
  //    G = Line Total
  //
  // 2) comma paste in A
  //    Description,Qty,Unit,Price
  /////////////////////////////////////
  if (supplierKey.indexOf('hazel') !== -1) {
    const firstCell = cell(0);
    const commaMapped = splitCommaRow_(firstCell);

    if (commaMapped && looksLikeMostlyBlankSplitRow_()) {
      return {
        itemCode: '',
        rawDesc: commaMapped[0] || '',
        packSize: '',
        pack: '',
        size: '',
        caseSize: '',
        casesOrdered: '',
        singlesOrdered: '',
        price: toNumber(commaMapped[3]),
        invoiceQty: toNumber(commaMapped[1]),
        unitsWeight: '',
        qtyType: commaMapped[2] || '',
        lineTotal: ''
      };
    }

    return {
      itemCode: cell(0),
      rawDesc: cell(1),
      packSize: '',
      pack: '',
      size: '',
      caseSize: '',
      casesOrdered: '',
      singlesOrdered: '',
      price: num(4),
      invoiceQty: num(3),
      unitsWeight: '',
      qtyType: cell(5),
      lineTotal: num(6)
    };
  }

    /////////////////////////////////////
  // BROADLAND HAMS
  // Supports:
  // 1) split columns
  //    A = Product Code
  //    B = Description
  //    C = Qty
  //    D = Unit
  //    E = Price
  //
  // 2) comma paste in column A
  //    Item Code,Description,Qty,Unit,Price
  /////////////////////////////////////
  if (supplierKey.indexOf('broadland') !== -1) {
    const firstCell = (rawRow[0] || '').toString().trim();

    const otherCellsUsed =
      ((rawRow[1] || '').toString().trim() !== '') ||
      ((rawRow[2] || '').toString().trim() !== '') ||
      ((rawRow[3] || '').toString().trim() !== '') ||
      ((rawRow[4] || '').toString().trim() !== '');

    if (firstCell.indexOf(',') !== -1 && !otherCellsUsed) {
      const parts = firstCell.split(',').map(p => (p || '').toString().trim());

      return {
        rawDesc: parts[1] || '',
        packSize: '',
        price: toNumber(parts[4]),
        invoiceQty: toNumber(parts[2]),
        unitsWeight: '',
        qtyType: parts[3] || '',
        itemCode: parts[0] || ''
      };
    }

    return {
      rawDesc: (rawRow[1] || '').toString().trim(),
      packSize: '',
      price: toNumber(rawRow[4]),
      invoiceQty: toNumber(rawRow[2]),
      unitsWeight: '',
      qtyType: (rawRow[3] || '').toString().trim(),
      itemCode: (rawRow[0] || '').toString().trim()
    };
  }

  /////////////////////////////////////
  // EASTERS
  // A = Description
  // B = Quantity
  // C = Unit
  // D = Unit Price
  /////////////////////////////////////
  if (supplierKey.indexOf('easter') !== -1) {
    const firstCell = cell(0);
    const commaMapped = splitCommaRow_(firstCell);

    // optional comma support without affecting normal split layout
    if (commaMapped && looksLikeMostlyBlankSplitRow_()) {
      return {
        itemCode: '',
        rawDesc: commaMapped[0] || '',
        packSize: '',
        pack: '',
        size: '',
        caseSize: '',
        casesOrdered: '',
        singlesOrdered: '',
        price: toNumber(commaMapped[3]),
        invoiceQty: toNumber(commaMapped[1]),
        unitsWeight: '',
        qtyType: commaMapped[2] || ''
      };
    }

    return {
      itemCode: '',
      rawDesc: cell(0),
      packSize: '',
      pack: '',
      size: '',
      caseSize: '',
      casesOrdered: '',
      singlesOrdered: '',
      price: num(3),
      invoiceQty: num(1),
      unitsWeight: '',
      qtyType: cell(2)
    };
  }

  /////////////////////////////////////
  // DEFAULT
  // Supports:
  // 1) split basic layout
  //    A = Description
  //    B = Price
  //    C = Quantity
  //    D = Item Code
  //
  // 2) comma paste in A
  /////////////////////////////////////
  const firstCell = cell(0);
  const genericComma = mapGenericCommaPaste_(firstCell);

  if (genericComma && looksLikeMostlyBlankSplitRow_()) {
    return genericComma;
  }

  return {
    itemCode: cell(3),
    rawDesc: cell(0),
    packSize: '',
    pack: '',
    size: '',
    caseSize: '',
    casesOrdered: '',
    singlesOrdered: '',
    price: num(1),
    invoiceQty: num(2),
    unitsWeight: '',
    qtyType: ''
  };
}

/////////////////////////////
// Clear Import Invoice Data
////////////////////////////
function clearInvoiceImport() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Invoice Import');
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert('Clear invoice?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const startRow = 8;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
  }

  sheet.getRange('B4').clearContent();
  sheet.getRange('B5').clearContent();
}

//////////////////////////////
// Build Invoice Import
/////////////////////////
function buildInvoiceImport() {
  const skippedRows = [];
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Import');
  const ui = SpreadsheetApp.getUi();

  try {
    if (!sheet) {
      ui.alert('Sheet "Invoice Import" not found.');
      return;
    }

    const context = getConfirmedInvoiceContext();
    if (!context) return;

    const supplier = context.supplier;
    const site = context.site;
    const supplierKey = (supplier || '').toString().trim().toLowerCase();

    const startRow = 8;
    const lastRawRow = getLastUsedRowInColumn(sheet, 1, startRow);

    if (lastRawRow < startRow) {
      ui.alert('No raw invoice rows found in column A from row 8 down.');
      return;
    }

    const rowCount = lastRawRow - startRow + 1;

    const rawData = sheet.getRange(startRow, 1, rowCount, 10).getValues();

    const existingGenerated = sheet.getRange(startRow, 4, rowCount, 25).getValues();

    const existingOverrideMap = buildExistingOverrideMap(rawData, existingGenerated);

    sheet.getRange(startRow, 4, rowCount, 25).clearContent(); // D:AB

    const rows = [];

    for (let i = 0; i < rawData.length; i++) {
      const mapped = getRawInvoiceFieldsBySupplier(rawData[i], supplier);

      const rawDesc = mapped.rawDesc;
      const packSize = mapped.packSize || mapped.pack || '';
      const sizeText = mapped.size || '';
      const caseSize = mapped.caseSize || '';
      const casesOrdered = mapped.casesOrdered;
      const singlesOrdered = mapped.singlesOrdered;
      const price = mapped.price;
      const invoiceQty = mapped.invoiceQty;
      const unitsWeight = mapped.unitsWeight || '';
      const qtyType = mapped.qtyType;
      const itemCode = mapped.itemCode || '';

      if (!rawDesc) continue;

      let parsed = parseInvoiceLineBySupplier(
        rawDesc,
        packSize,
        price,
        invoiceQty,
        supplier,
        qtyType,
        unitsWeight,
        sizeText,
        caseSize,
        casesOrdered,
        singlesOrdered
      );

      /////////////////////////////////////
      // BIDFOOD FIX
      /////////////////////////////////////
      if (supplierKey.indexOf('bidfood') !== -1) {
        const packParsed = parseBidfoodPackSize(
          packSize,
          price,
          invoiceQty,
          unitsWeight
        );

        if (packParsed) {
          parsed.packSizeDisplay = packParsed.packSizeDisplay;
          parsed.packQty = packParsed.packQty;
          parsed.baseUnit = packParsed.baseUnit;
          parsed.effectivePackPrice = packParsed.effectivePackPrice;

          if (packParsed.effectivePackPrice && packParsed.packQty) {
            parsed.costPerUnit = roundNumber(
              packParsed.effectivePackPrice / Number(packParsed.packQty),
              6
            );
          }
        }
      }

      let reviewFlag = getEnhancedReviewFlag(parsed, rawDesc, price, invoiceQty);

      if (Number(invoiceQty) < 0) {
        reviewFlag = 'CREDIT';
      }

      const rowKey = makeImportRowKey(rawDesc, price, invoiceQty, itemCode);
      const existing = existingOverrideMap[rowKey] || null;

      const defaultIngredient = (parsed.suggestedIngredient || '').toString().trim();
      const defaultPackSize = parsed.packSizeDisplay || '';
      const defaultPackQty = parsed.packQty;
      const defaultBaseUnit = (parsed.baseUnit || '').toString().trim();
      const defaultCostPerUnit = parsed.costPerUnit;

      const overrideIngredient = existing ? existing.overrideIngredient : '';
      const overridePackSize = existing ? existing.overridePackSize : '';
      const overridePackQty = existing ? existing.overridePackQty : '';
      const overrideBaseUnit = existing ? existing.overrideBaseUnit : '';

      const finalIngredient = overrideIngredient || defaultIngredient;
      const finalPackSize = overridePackSize || defaultPackSize;
      const finalPackQty = overridePackQty !== '' ? overridePackQty : defaultPackQty;
      const finalBaseUnit = overrideBaseUnit || defaultBaseUnit;

      let finalCostPerUnit = '';
      if (parsed.effectivePackPrice && finalPackQty !== '' && Number(finalPackQty) !== 0) {
        finalCostPerUnit = roundNumber(parsed.effectivePackPrice / Number(finalPackQty), 6);
      }

      let finalCleanName = finalIngredient
        ? finalIngredient.toString().trim().toLowerCase()
        : '';

      // THYME FIX
      finalCleanName = finalCleanName
        .replace(/\bff thyme\b/gi, 'fresh thyme');

      // NORMALISATION
      finalCleanName = finalCleanName
        .replace(/\bliptons\b/g, 'lipton')
        .replace(/\b\d+\s*s\b/g, '')
        .replace(/\b\d+s\b/g, '')
        .replace(/\bpet\b/g, '')
        .replace(/\b\d+\s*(ml|l|g|kg)\b/g, '')
        .replace(/\s+/g, ' ')
        .trim();

      finalCleanName = finalCleanName
        .replace(/\bliptons\b/g, 'lipton')
        .replace(/\b\d+s\b/g, '')
        .replace(/\bpet\b/g, '')
        .replace(/\s+/g, ' ')
        .trim();

      let finalReviewFlag = getFinalReviewFlag(
        finalIngredient,
        finalPackSize,
        finalPackQty,
        finalBaseUnit,
        reviewFlag
      );

      if (Number(invoiceQty) < 0) {
        finalReviewFlag = 'CREDIT';
      }

      // VALIDATION (skip bad rows)
      if (!finalPackQty || !parsed.effectivePackPrice) {
        skippedRows.push({
          raw: rawDesc,
          reason: 'Missing packQty or price'
        });
        continue;
      }

      if (!finalBaseUnit) {
        skippedRows.push({
          raw: rawDesc,
          reason: 'Missing base unit'
        });
        continue;
      }

      rows.push({
        main: [
          defaultIngredient,                              // D
          overrideIngredient,                             // E
          finalIngredient,                                // F
          finalCleanName,                                 // G
          defaultPackSize,                                // H
          overridePackSize,                               // I
          finalPackSize,                                  // J
          defaultPackQty,                                 // K
          overridePackQty,                                // L
          finalPackQty,                                   // M
          defaultBaseUnit,                                // N
          overrideBaseUnit,                               // O
          finalBaseUnit,                                  // P
          finalCostPerUnit || defaultCostPerUnit || '',   // Q
          finalReviewFlag                                 // R
        ],
        master: [
          finalIngredient,                                // T
          finalCleanName,                                 // U
          supplier,                                       // V
          finalPackSize,                                  // W
          finalPackQty,                                   // X
          parsed.effectivePackPrice || '',                // Y
          finalBaseUnit,                                  // Z
          finalCostPerUnit || defaultCostPerUnit || '',   // AA
          itemCode                                        // AB
        ],
        sortFlag: finalReviewFlag === 'OK' ? 1 : 0,
        ingredient: finalIngredient ? finalIngredient.toString().toLowerCase() : ''
      });
    }

    rows.sort((a, b) => {
      if (a.sortFlag !== b.sortFlag) return a.sortFlag - b.sortFlag;
      return a.ingredient.localeCompare(b.ingredient);
    });

    const outputMain = rows.map(r => r.main);
    const outputMaster = rows.map(r => r.master);

    if (outputMain.length === 0) {
      ui.alert('No usable raw invoice rows were found.');
      return;
    }

    sheet.getRange(startRow, 4, outputMain.length, 15).setValues(outputMain);     // D:R
    sheet.getRange(startRow, 20, outputMaster.length, 9).setValues(outputMaster); // T:AB

    sheet.getRange(startRow, 17, outputMain.length, 1).setNumberFormat('£0.0000');   // Q
    sheet.getRange(startRow, 25, outputMaster.length, 1).setNumberFormat('£0.00');   // Y
    sheet.getRange(startRow, 27, outputMaster.length, 1).setNumberFormat('£0.0000'); // AA

    if (skippedRows.length > 0) {
      const grouped = {};

      skippedRows.forEach(r => {
        if (!grouped[r.reason]) grouped[r.reason] = [];
        grouped[r.reason].push(r.raw);
      });

      let message =
        `Build complete with skips\n\n` +
        `Total Skipped: ${skippedRows.length}\n\n`;

      Object.keys(grouped).forEach(reason => {
        message += `⚠ ${reason} (${grouped[reason].length})\n`;

        grouped[reason].slice(0, 5).forEach(item => {
          message += `• ${item}\n`;
        });

        if (grouped[reason].length > 5) {
          message += `...and ${grouped[reason].length - 5} more\n`;
        }

        message += '\n';
      });

      ui.alert(message);
    }

    ui.alert(
      `Invoice import built for ${supplier}.\n\nRows processed: ${rows.length}\nManual overrides preserved where possible.`
    );

  } catch (err) {
    ui.alert(
      'buildInvoiceImport error:\n\n' + (err && err.message ? err.message : err)
    );
    throw err;
  }
}
//////////////////////////////////////////
// BUILD + APPEND INVOICE IMPORT
// SAFE WRAPPER ONLY
//////////////////////////////////////////
function buildAndAppendInvoiceImport() {
  if (!requirePdfReviewComplete_()) return;
  const ss = SpreadsheetApp.getActive();
  const importSheet = ss.getSheetByName('Invoice Import');
  const ui = SpreadsheetApp.getUi();

  try {
    if (!importSheet) {
      ui.alert('Missing "Invoice Import" sheet.');
      return;
    }

    /////////////////////////////////////////
    // STEP 1: BUILD
    /////////////////////////////////////////
    buildInvoiceImport();

    /////////////////////////////////////////
    // CLEAR OLD HIGHLIGHTS
    /////////////////////////////////////////
    clearInvoiceImportReviewHighlights_();

    const startRow = 8;
    const lastRow = getLastUsedRowInColumn(importSheet, 4, startRow); // built output starts at D

    if (lastRow < startRow) {
      ui.alert(
        'Build completed, but no rows were built.\n\n' +
        'Nothing has been appended to Ingredients Master.'
      );
      return;
    }

    const rowCount = lastRow - startRow + 1;
    const finalBlock = importSheet.getRange(startRow, 4, rowCount, 15).getValues(); // D:R

    let builtCount = 0;
    let criticalSkips = [];
    let warningSkips = 0;
    let suspiciousPackRows = [];

    const criticalRowNumbers = [];
    const suspiciousRowNumbers = [];

    for (let i = 0; i < finalBlock.length; i++) {
      const rowNumber = startRow + i;

      const ingredient = (finalBlock[i][2] || '').toString().trim();   // F
      const packSizeText = (finalBlock[i][6] || '').toString().trim(); // J
      const packQty = finalBlock[i][9];                                // M
      const baseUnit = (finalBlock[i][12] || '').toString().trim();    // P
      const reviewFlag = (finalBlock[i][14] || '').toString().trim().toUpperCase(); // R

      if (ingredient) builtCount++;

      const isSuspicious = ingredient &&
        isSuspiciousPackSize_(ingredient, packQty, baseUnit, packSizeText);

      if (!reviewFlag) {
        if (isSuspicious) {
          suspiciousPackRows.push(
            `Row ${rowNumber}: ${ingredient} | ${packSizeText} | Qty ${packQty} ${baseUnit}`
          );
          suspiciousRowNumbers.push(rowNumber);
        }
        continue;
      }

      if (
        reviewFlag === 'SKIP' ||
        reviewFlag === 'CHECK' ||
        reviewFlag === 'CHECK PRICE' ||
        reviewFlag === 'CHECK PACK' ||
        reviewFlag === 'CHECK WEIGHT' ||
        reviewFlag === 'CHECK UNIT' ||
        reviewFlag === 'CHECK QTY'
      ) {
        criticalSkips.push(`Row ${rowNumber}: ${reviewFlag}`);
        criticalRowNumbers.push(rowNumber);
      } else if (reviewFlag === 'CREDIT') {
        warningSkips++;
      }

      if (isSuspicious) {
        suspiciousPackRows.push(
          `Row ${rowNumber}: ${ingredient} | ${packSizeText} | Qty ${packQty} ${baseUnit}`
        );
        suspiciousRowNumbers.push(rowNumber);
      }
    }

    /////////////////////////////////////////
    // SAFETY CHECKS
    /////////////////////////////////////////
    if (builtCount === 0) {
      ui.alert(
        'Build completed, but no valid built rows were found.\n\n' +
        'Nothing has been appended to Ingredients Master.'
      );
      return;
    }

    if (criticalSkips.length > 0) {
      highlightInvoiceImportRows_(criticalRowNumbers, '#f4cccc'); // light red

      let message =
        'Build stopped before append.\n\n' +
        'Critical review rows were found in Invoice Import.\n\n' +
        criticalSkips.slice(0, 15).join('\n');

      if (criticalSkips.length > 15) {
        message += `\n...and ${criticalSkips.length - 15} more`;
      }

      message += '\n\nThese rows have been highlighted in red.';

      ui.alert(message);
      importSheet.getRange(criticalRowNumbers[0], 4).activate();
      return;
    }

    if (suspiciousPackRows.length > 0) {
      highlightInvoiceImportRows_(suspiciousRowNumbers, '#fff2cc'); // light amber

      let message =
        'Build stopped before append.\n\n' +
        'Suspicious large pack sizes were found.\n\n' +
        suspiciousPackRows.slice(0, 15).join('\n');

      if (suspiciousPackRows.length > 15) {
        message += `\n...and ${suspiciousPackRows.length - 15} more`;
      }

      message += '\n\nThese rows have been highlighted in amber.';

      ui.alert(message);
      importSheet.getRange(suspiciousRowNumbers[0], 4).activate();
      return;
    }

    /////////////////////////////////////////
    // STEP 2: APPEND
    /////////////////////////////////////////
    appendInvoiceRowsToIngredientsMaster();

  } catch (err) {
    ui.alert(
      'buildAndAppendInvoiceImport error:\n\n' +
      (err && err.message ? err.message : err)
    );
    throw err;
  }
}