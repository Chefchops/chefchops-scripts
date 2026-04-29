/////////////////////////////////////
// EASTERS PARSER
// Rule: Unit Price is source of truth
/////////////////////////////////////
function parseEastersLine(rawDesc, packPrice, invoiceQty, qtyType) {
  const original = (rawDesc || '').toString().trim();
  let cleaned = original.replace(/\s+/g, ' ').trim();

  const effectivePackPrice = toNumber(packPrice); // UNIT PRICE ONLY
  const unitText = (qtyType || '').toString().trim().toLowerCase();

  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let reviewFlag = 'OK';

  cleaned = cleaned
    .replace(/\(\s*per kg\s*\)/ig, '')
    .replace(/\(\s*tray\s*\)/ig, '')
    .replace(/\s+/g, ' ')
    .trim();
/////////////////////////////////////
// HANDLE BOX + WEIGHT (e.g. 1.3Kg)
/////////////////////////////////////

let weightMatch = cleaned.match(/(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr)/i);

if (weightMatch) {
  const qty = toNumber(weightMatch[1]);
  const unit = weightMatch[2].toLowerCase();

  if (unit === 'kg') {
    packQty = qty * 1000;
    baseUnit = 'g';
  } else if (unit === 'g') {
    packQty = qty;
    baseUnit = 'g';
  } else if (unit === 'l' || unit === 'ltr') {
    packQty = qty * 1000;
    baseUnit = 'ml';
  } else if (unit === 'ml') {
    packQty = qty;
    baseUnit = 'ml';
  }

  packSizeDisplay = weightMatch[0];
}
  /////////////////////////////////////
  // FIXED SPECIAL CASE
  /////////////////////////////////////
  if (/20\s*x\s*125\s*g/i.test(original)) {
    return {
      original,
      suggestedIngredient: 'Yoghurt Thick&& Creamy Mixed Fruit',
      cleanName: 'yoghurt thick&& creamy mixed fruit',
      packSizeDisplay: '20x125g',
      packQty: 20,
      baseUnit: 'unit',
      effectivePackPrice: toNumber(packPrice),
      costPerUnit: roundNumber(toNumber(packPrice) / 20, 6),
      reviewFlag: 'OK'
    };
  }

  /////////////////////////////////////
  // EGGS (TRAY HANDLING LIKE HAZELS)
  /////////////////////////////////////
  let eggMatch = original.match(/(\d+)\s*(tray|trays)\s*(of)?\s*(\d+)/i);

  if (eggMatch) {
    const trays = Number(eggMatch[1]);
    const eggsPerTray = Number(eggMatch[4]);
    const totalEggs = trays * eggsPerTray;

    packSizeDisplay = `${trays} tray of ${eggsPerTray}`;
    packQty = totalEggs;
    baseUnit = 'unit';

    const suggestedIngredient = toTitleCase(
      original
        .replace(eggMatch[0], '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // EGGS FALLBACK (TRAY 30 FORMAT)
  /////////////////////////////////////
  let eggSimple = original.match(/tray\s*(\d+)/i);

  if (eggSimple && /egg/i.test(original)) {
    const eggs = Number(eggSimple[1]);

    packSizeDisplay = `1 tray of ${eggs}`;
    packQty = eggs;
    baseUnit = 'unit';

    const suggestedIngredient = toTitleCase(
      original
        .replace(eggSimple[0], '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // MULTI-PACK LIKE 20x125g
  /////////////////////////////////////
  let m = cleaned.match(/(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);
  if (m) {
    const outerQty = Number(m[1]);
    const innerQty = Number(m[2]);
    const unit = normaliseUnit(m[3]);

    packSizeDisplay = `${outerQty}x${innerQty}${unit}`;
    packQty = outerQty;
    baseUnit = 'unit';

    const suggestedIngredient = toTitleCase(
      original
        .replace(m[0], '')
        .replace(/\(\s*tray\s*\)/ig, '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // SINGLE PACK SIZE IN DESCRIPTION
  /////////////////////////////////////
  m = cleaned.match(/(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);
  if (m) {
    const size = Number(m[1]);
    const unit = normaliseUnit(m[2]);

    packSizeDisplay = `${stripTrailingZeros_(size)}${unit}`;
    packQty = convertToBaseQty(size, unit);
    baseUnit = unitToBaseUnit(unit);

    const suggestedIngredient = toTitleCase(
      original
        .replace(m[0], '')
        .replace(/\bbox\b/ig, '')
        .replace(/\bcase\b/ig, '')
        .replace(/\(\s*tray\s*\)/ig, '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // SOLD PER KG
  /////////////////////////////////////
  if (/\bper\s*kg\b/i.test(original) || unitText === 'kg') {
    packSizeDisplay = 'per kg';
    packQty = 1000;
    baseUnit = 'g';

    const suggestedIngredient = toTitleCase(
      original
        .replace(/\(\s*per kg\s*\)/ig, '')
        .replace(/\bsingle\b/ig, '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // SINGLE / LOAF / EACH ITEMS
  /////////////////////////////////////
  if (
    unitText === 'single' ||
    /\bsingle\b/i.test(original) ||
    /\bloaf\b/i.test(original)
  ) {
    packSizeDisplay = 'unit';
    packQty = 1;
    baseUnit = 'unit';

    const suggestedIngredient = toTitleCase(
      original
        .replace(/\bsingle\b/ig, '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // BOX / CASE WITHOUT READABLE SIZE
  /////////////////////////////////////
  if (unitText === 'box' || unitText === 'case') {
    packSizeDisplay = `1${unitText}`;
    packQty = 1;
    baseUnit = 'unit';
    reviewFlag = 'CHECK UNIT';

    const suggestedIngredient = toTitleCase(
      original
        .replace(/\bbox\b/ig, '')
        .replace(/\bcase\b/ig, '')
        .replace(/\(\s*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );

    const cleanName = cleanIngredientName_(suggestedIngredient);

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

    return {
      original,
      suggestedIngredient,
      cleanName,
      packSizeDisplay,
      packQty,
      baseUnit,
      effectivePackPrice,
      costPerUnit,
      reviewFlag
    };
  }

  /////////////////////////////////////
  // FALLBACK
  /////////////////////////////////////
  return parseInvoiceLine(rawDesc, packPrice, invoiceQty, qtyType);
}