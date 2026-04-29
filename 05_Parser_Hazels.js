/////////////////////////////////////
// HAZELS PARSER
/////////////////////////////////////
function parseHazelsLine(rawDesc, packPrice, invoiceQty, qtyType) {
  const original = (rawDesc || '').toString().trim();
  const suggestedIngredient = toTitleCase(original);
  const cleanName = cleanIngredientName_(suggestedIngredient);

  const qty = toNumber(invoiceQty);
  const unit = (qtyType || '').toString().trim().toLowerCase();
  const price = toNumber(packPrice);

  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let effectivePackPrice = '';
  let costPerUnit = '';
  let reviewFlag = 'OK';

  const isEachUnit = unit === 'each' || unit === 'ea';
  const isPackUnit = unit === 'pack' || unit === 'pk';

  const looksLikeButcherWeightItem =
    cleanName.includes('case') ||
    cleanName.includes('fillet') ||
    cleanName.includes('fillets') ||
    cleanName.includes('breast') ||
    cleanName.includes('chicken') ||
    cleanName.includes('beef') ||
    cleanName.includes('pork') ||
    cleanName.includes('loin') ||
    cleanName.includes('joint') ||
    cleanName.includes('steak') ||
    cleanName.includes('mince') ||
    cleanName.includes('sausage') ||
    cleanName.includes('sausages');

  /////////////////////////////////////
  // EGGS
  // Hazels eggs are treated as trays of 30
  /////////////////////////////////////
  if (cleanName.includes('egg')) {
    const traySize = 30;

    packSizeDisplay = '30ea';
    packQty = traySize;
    baseUnit = 'unit';

    effectivePackPrice = price;
    costPerUnit = price > 0 ? roundNumber(price / traySize, 6) : '';

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
  // BACON CASE RULE
  // Bacon cases are treated as 4 packs
  /////////////////////////////////////
  if (cleanName.includes('bacon') && isEachUnit && price > 0) {
    packSizeDisplay = '4ea';
    packQty = 4;
    baseUnit = 'unit';

    effectivePackPrice = price;
    costPerUnit = roundNumber(price / 4, 6);

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
  // PER KG ITEMS
  /////////////////////////////////////
  if (unit === 'kg' && price > 0) {
    packSizeDisplay = '1kg';
    packQty = 1000;
    baseUnit = 'g';

    effectivePackPrice = price;
    costPerUnit = roundNumber(price / 1000, 6);

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
  // PER GRAM ITEMS
  /////////////////////////////////////
  if (unit === 'g' && price > 0) {
    packSizeDisplay = '1g';
    packQty = 1;
    baseUnit = 'g';

    effectivePackPrice = price;
    costPerUnit = price;

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
  // SUSPICIOUS BUTCHER "EA" LINES
  /////////////////////////////////////
  if (isEachUnit && price > 0 && looksLikeButcherWeightItem) {
    packSizeDisplay = '1ea';
    packQty = 1;
    baseUnit = 'unit';

    effectivePackPrice = price;
    costPerUnit = price;
    reviewFlag = 'CHECK WEIGHT';

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
  // GENUINE EACH ITEMS
  /////////////////////////////////////
  if (isEachUnit && price > 0) {
    packSizeDisplay = '1ea';
    packQty = 1;
    baseUnit = 'unit';

    effectivePackPrice = price;
    costPerUnit = price;

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
  // PER PACK ITEMS
  /////////////////////////////////////
  if (isPackUnit && price > 0) {
    packSizeDisplay = '1pk';
    packQty = 1;
    baseUnit = 'unit';

    effectivePackPrice = price;
    costPerUnit = price;

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
  reviewFlag = 'CHECK UNIT';

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