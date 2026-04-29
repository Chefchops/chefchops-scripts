/////////////////////////////////////
// GENERIC PARSER
/////////////////////////////////////
function parseInvoiceLine(rawDesc, packPrice, invoiceQty, qtyType) {
  let text = (rawDesc || '').toString().trim();
  const original = text;

  let cleaned = text
    .replace(/\([^)]*\)/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let reviewFlag = 'OK';
  let effectivePackPrice = packPrice;
  let isWeightedPerKg = false;

  let m;

  /////////////////////////////////////
  // MULTIPACK WITH WEIGHT / VOLUME
  // Example: 20x125g
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);
    if (m) {
      const outerQty = Number(m[1]);
      const innerQty = Number(m[2]);
      const unit = normaliseUnit(m[3]);

      packSizeDisplay = `${outerQty}x${innerQty}${unit}`;
      baseUnit = unitToBaseUnit(unit);
      packQty = convertToBaseQty(outerQty * innerQty, unit);

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\b(tray|box|case|pack)\b/ig, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // DASH MULTIPACK WITH WEIGHT / VOLUME
  // Example: 6-850g
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);
    if (m) {
      const outerQty = Number(m[1]);
      const innerQty = Number(m[2]);
      const unit = normaliseUnit(m[3]);

      packSizeDisplay = `${outerQty}x${innerQty}${unit}`;
      baseUnit = unitToBaseUnit(unit);
      packQty = convertToBaseQty(outerQty * innerQty, unit);

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\b(tray|box|case|pack)\b/ig, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // SINGLE PACK WEIGHT / VOLUME
  // Example: 1kg / 500ml
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);
    if (m) {
      const qty = Number(m[1]);
      const unit = normaliseUnit(m[2]);

      packSizeDisplay = `${qty}${unit}`;
      baseUnit = unitToBaseUnit(unit);
      packQty = convertToBaseQty(qty, unit);

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\b(box|case|pack)\b/ig, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // SIMPLE MULTIPACK COUNT
  // Example: 4x6
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+)\s*x\s*(\d+)\b/i);
    if (m) {
      const outerQty = Number(m[1]);
      const innerQty = Number(m[2]);

      packSizeDisplay = `${outerQty}x${innerQty}`;
      packQty = outerQty * innerQty;
      baseUnit = 'each';

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\b(tray|box|case|pack)\b/ig, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // DASH COUNT FORMAT
  // Example: 4-24-...
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+)\s*-\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\b/);
    if (m) {
      const caseQty = Number(m[1]);

      packSizeDisplay = `${caseQty}each`;
      packQty = caseQty;
      baseUnit = 'each';

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // DASH COUNT WITH UNIT LABEL
  // Example: 4-25pk
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+)\s*-\s*(\d+(?:\.\d+)?)\s*(pk|ea|ptn|portion|portions|sti|sticks|roll|rolls|can|cans|btl|btls|unit|units|sac|sachets?)\b/i);
    if (m) {
      const outerQty = Number(m[1]);
      const innerQty = Number(m[2]);
      const unitLabel = m[3].toLowerCase();

      packSizeDisplay = `${outerQty}x${innerQty}${unitLabel}`;
      packQty = outerQty * innerQty;
      baseUnit = 'each';

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // SIMPLE COUNT WITH UNIT LABEL
  // Example: 30ea / 12pk
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\b(\d+(?:\.\d+)?)\s*(pk|ea|ptn|portion|portions|sti|sticks|roll|rolls|can|cans|btl|btls|unit|units|sac|sachets?)\b/i);
    if (m) {
      const qty = Number(m[1]);

      packSizeDisplay = `${qty}${m[2].toLowerCase()}`;
      packQty = qty;
      baseUnit = 'each';

      cleaned = cleaned
        .replace(m[0], '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // SOLD PER KG
  /////////////////////////////////////
  if (!packQty && /per\s*kg/i.test(text)) {
    packSizeDisplay = 'per kg';
    packQty = 1000;
    baseUnit = 'g';
    isWeightedPerKg = true;

    cleaned = cleaned
      .replace(/single/ig, '')
      .replace(/per\s*kg/ig, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /////////////////////////////////////
  // BACON RULE
  // Bacon cases are treated as 4 packs across suppliers
  /////////////////////////////////////
  if (
    !packQty &&
    /\bbacon\b/i.test(text) &&
    (
      /\bsingle\b/i.test(text) ||
      /\bea\b/i.test((qtyType || '').toString()) ||
      /\beach\b/i.test((qtyType || '').toString()) ||
      /\bunit\b/i.test((qtyType || '').toString())
    )
  ) {
    packSizeDisplay = '4ea';
    packQty = 4;
    baseUnit = 'unit';

    cleaned = cleaned
      .replace(/single/ig, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /////////////////////////////////////
  // SINGLE ITEM
  /////////////////////////////////////
  if (!packQty && /\bsingle\b/i.test(text)) {
    packSizeDisplay = 'each';
    packQty = 1;
    baseUnit = 'each';

    cleaned = cleaned
      .replace(/single/ig, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /////////////////////////////////////
  // BREAD / BAKERY SINGLE ITEM FALLBACK
  /////////////////////////////////////
  if (!packQty && /\b(loaf|bloomer|tin|bap|bun|baguette|teacake|muffin|crumpet|pitta|naan|wrap)\b/i.test(cleaned)) {
    packSizeDisplay = 'each';
    packQty = 1;
    baseUnit = 'each';
  }

  /////////////////////////////////////
  // BOX SIZE FORMAT
  // Example: Size 40 box
  /////////////////////////////////////
  if (!packQty) {
    m = text.match(/\bsize\s*(\d+)\b/i);
    if (m && /\bbox\b/i.test(text)) {
      packSizeDisplay = `box of ${m[1]}`;
      packQty = Number(m[1]);
      baseUnit = 'each';

      cleaned = cleaned
        .replace(/\bbox\b/ig, '')
        .replace(/\s+/g, ' ')
        .trim();
    }
  }

  /////////////////////////////////////
  // WEIGHTED FALLBACK USING INVOICE QTY
  /////////////////////////////////////
  if (!packQty && invoiceQty && Number(invoiceQty) > 0) {
    packSizeDisplay = 'per kg';
    packQty = 1000;
    baseUnit = 'g';
    isWeightedPerKg = true;
  }

  /////////////////////////////////////
  // BOX / CASE / PACK FALLBACK
  /////////////////////////////////////
  if (!packQty && /\b(box|case|pack|pk)\b/i.test(text)) {
    packSizeDisplay = 'box';
    packQty = '';
    baseUnit = '';
    reviewFlag = 'CHECK PACK';

    cleaned = cleaned
      .replace(/\bbox\b|\bcase\b|\bpack\b|\bpk\b/ig, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /////////////////////////////////////
  // FINAL NAME CLEANUP
  /////////////////////////////////////
  cleaned = cleaned
    .replace(/\bSPECIAL PRICE \d+\b/ig, '')
    .replace(/\bSPECIAL PRICE\b/ig, '')
    .replace(/\bPRICE \d+\b/ig, '')
    .replace(/\bL\b/g, ' ')
    .replace(/\bCL\b/g, ' ')
    .replace(/\bG\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  let finalName = cleaned.trim();

  if (/^bread\s+/i.test(finalName)) {
    finalName = finalName.replace(/^bread\s+/i, '');
    finalName = finalName.replace(/\bloaf\b/i, 'bread loaf');
  }

  finalName = finalName
    .replace(/\bbread loaf bread\b/i, 'bread loaf')
    .replace(/\s+/g, ' ')
    .trim();

  let ingredient = toTitleCase(finalName);
  const cleanName = ingredient.toLowerCase().trim();

  /////////////////////////////////////
  // EFFECTIVE PACK PRICE FOR WEIGHTED ITEMS
  /////////////////////////////////////
  if (isWeightedPerKg && invoiceQty && Number(invoiceQty) > 0) {
    effectivePackPrice = roundNumber(packPrice / Number(invoiceQty), 2);
  }

  /////////////////////////////////////
  // PRICE CHECK
  /////////////////////////////////////
  if (isNaN(Number(effectivePackPrice))) {
    reviewFlag = 'CHECK PRICE';
  }

  /////////////////////////////////////
  // COST PER UNIT
  /////////////////////////////////////
  let costPerUnit = '';
  if (!isNaN(effectivePackPrice) && Number(packQty) > 0) {
    costPerUnit = roundNumber(effectivePackPrice / Number(packQty), 6);
  }

  /////////////////////////////////////
  // FINAL REVIEW CHECK
  /////////////////////////////////////
  if (!ingredient || !packQty || !baseUnit) {
    reviewFlag = reviewFlag === 'OK' ? 'CHECK' : reviewFlag;
  }

  return {
    original,
    suggestedIngredient: ingredient,
    cleanName,
    packSizeDisplay,
    packQty,
    baseUnit,
    effectivePackPrice,
    costPerUnit,
    reviewFlag
  };
}