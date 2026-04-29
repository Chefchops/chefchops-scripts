/////////////////////////////////////
// PILGRIM PARSER
/////////////////////////////////////
function parsePilgrimLine(rawDesc, caseSizeText, packSizeText, unitPrice, casesOrdered, singlesOrdered) {
  let text = (rawDesc || '').toString().trim();
  let caseRaw = (caseSizeText || '').toString().trim();
  let sizeRaw = (packSizeText || '').toString().trim();

  const upper = text.toUpperCase();

  if (!text || upper.includes('VOID') || upper.includes('CREDIT')) {
    return {
      suggestedIngredient: '',
      packSizeDisplay: '',
      packQty: '',
      baseUnit: '',
      effectivePackPrice: '',
      costPerUnit: '',
      reviewFlag: 'SKIP',
      unitCount: ''
    };
  }

  let cleaned = text.replace(/\s+/g, ' ').trim();

  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let unitCount = '';
  let reviewFlag = 'OK';

  let caseSize = toNumber(caseRaw);

  if (caseRaw && /(dozen|doz|dz)/i.test(caseRaw)) {
    let dm = caseRaw.match(/(\d+(?:\.\d+)?)/);
    if (dm) {
      caseSize = Number(dm[1]) * 12;
    }
  }

  if (isNaN(caseSize) || caseSize <= 0) caseSize = 1;

  let casesQty = toNumber(casesOrdered);
  if (isNaN(casesQty) || casesQty < 0) casesQty = 0;

  let singlesQty = toNumber(singlesOrdered);
  if (isNaN(singlesQty) || singlesQty < 0) singlesQty = 0;

  let pricedUnitMultiplier = casesQty > 0 ? caseSize : 1;

  let m;
  let qtyPerInnerPack = '';

  sizeRaw = sizeRaw
    .replace(/\s*[xX]\s*/g, 'x')
    .replace(/\s+/g, ' ')
    .trim();

  function normalisePilgrimCountUnit_(unitRaw) {
    const u = (unitRaw || '').toString().trim().toLowerCase();
    if (u === 'pk') return 'pk';
    if (u === 'ea' || u === 'unit' || u === 'units') return 'ea';
    if (u === 'sti' || u === 'stick' || u === 'sticks') return 'stick';
    if (u === 'ptn' || u === 'portion' || u === 'portions') return 'portion';
    return 'unit';
  }

  /////////////////////////////////////
  // MULTIPACK WEIGHT / VOLUME / LENGTH
  // e.g. 4x2.27kg / 2x1kg / 12x500ml / 50x135g / 24x35g
  /////////////////////////////////////
  if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter|m)$/i))) {
    const outerCount = Number(m[1]);
    const innerQty = Number(m[2]);
    const unitRaw = m[3];
    const unit = normaliseUnit(unitRaw);

    qtyPerInnerPack = convertToBaseQty(innerQty, unit) * outerCount;

    baseUnit = normaliseUnit(unit);
    if (baseUnit === 'kg') baseUnit = 'g';
    if (baseUnit === 'l') baseUnit = 'ml';

    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);

    if (pricedUnitMultiplier > 1) {
      packSizeDisplay = `${pricedUnitMultiplier}x${outerCount}x${innerQty}${unitRaw}`;
    } else {
      packSizeDisplay = `${outerCount}x${innerQty}${unitRaw}`;
    }
  }

  /////////////////////////////////////
  // MULTIPACK COUNT UNITS
  // e.g. 10x100ea / 4x50pk / 2x1000sti
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)\s*(pk|ea|unit|units|sti|stick|sticks|ptn|portion|portions)$/i))) {
    const outerCount = Number(m[1]);
    const innerQty = Number(m[2]);
    const unitLabel = normalisePilgrimCountUnit_(m[3]);

    baseUnit = 'unit';
    qtyPerInnerPack = outerCount * innerQty;
    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);
    unitCount = packQty;

    if (pricedUnitMultiplier > 1) {
      packSizeDisplay = `${pricedUnitMultiplier}x${outerCount}x${innerQty}${unitLabel}`;
    } else {
      packSizeDisplay = `${outerCount}x${innerQty}${unitLabel}`;
    }
  }

  /////////////////////////////////////
  // SINGLE WEIGHT / VOLUME / LENGTH
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter|m)$/i))) {
    const innerQty = Number(m[1]);
    const unitRaw = m[2];
    const unit = normaliseUnit(unitRaw);

    qtyPerInnerPack = convertToBaseQty(innerQty, unit);

    baseUnit = normaliseUnit(unit);
    if (baseUnit === 'kg') baseUnit = 'g';
    if (baseUnit === 'l') baseUnit = 'ml';

    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);

    if (pricedUnitMultiplier > 1) {
      packSizeDisplay = `${pricedUnitMultiplier}x${innerQty}${unitRaw}`;
    } else {
      packSizeDisplay = `${innerQty}${unitRaw}`;
    }
  }

  /////////////////////////////////////
  // DOZEN / DOZ / DZ IN PACK SIZE
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)?\s*(dozen|doz|dz)$/i))) {
    const innerQty = m[1] ? Number(m[1]) * 12 : 12;

    baseUnit = 'unit';
    qtyPerInnerPack = innerQty;
    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);
    unitCount = packQty;

    if (pricedUnitMultiplier > 1) {
      packSizeDisplay = `${pricedUnitMultiplier}x${innerQty}`;
    } else {
      packSizeDisplay = `${innerQty}`;
    }
  }

  /////////////////////////////////////
  // SINGLE COUNT ITEMS
  // e.g. 30s / 60pk / 1000sti / 18ptn / 180pk
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)\s*(s|pk|ea|unit|units|sti|stick|sticks|ptn|portion|portions)$/i))) {
    const innerQty = Number(m[1]);
    const rawUnit = m[2].toLowerCase();

    baseUnit = 'unit';
    qtyPerInnerPack = innerQty;
    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);
    unitCount = packQty;

    if (rawUnit === 's') {
      packSizeDisplay = pricedUnitMultiplier > 1 ? `${pricedUnitMultiplier}x${innerQty}s` : `${innerQty}s`;
    } else {
      const unitLabel = normalisePilgrimCountUnit_(rawUnit);
      packSizeDisplay = pricedUnitMultiplier > 1 ? `${pricedUnitMultiplier}x${innerQty}${unitLabel}` : `${innerQty}${unitLabel}`;
    }
  }

  /////////////////////////////////////
  // SINGLE PACK COUNT LIKE 972 OR 1 OR 10
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)$/i))) {
    const innerQty = Number(m[1]);

    baseUnit = 'unit';
    qtyPerInnerPack = innerQty;
    packQty = roundNumber(qtyPerInnerPack * pricedUnitMultiplier, 4);
    unitCount = packQty;

    if (pricedUnitMultiplier > 1) {
      packSizeDisplay = `${pricedUnitMultiplier}x${innerQty}`;
    } else {
      packSizeDisplay = `${innerQty}`;
    }
  }

  /////////////////////////////////////
  // FALLBACK: use case size only if pack size blank
  /////////////////////////////////////
  else if (!sizeRaw && caseRaw) {
    if ((m = caseRaw.match(/^(\d+(?:\.\d+)?)\s*(dozen|doz|dz)$/i))) {
      const innerQty = Number(m[1]) * 12;
      baseUnit = 'unit';
      qtyPerInnerPack = innerQty;
      packQty = innerQty;
      unitCount = innerQty;
      packSizeDisplay = `${innerQty}`;
    } else if ((m = caseRaw.match(/^(\d+(?:\.\d+)?)$/i))) {
      const innerQty = Number(m[1]);
      baseUnit = 'unit';
      qtyPerInnerPack = innerQty;
      packQty = innerQty;
      unitCount = innerQty;
      packSizeDisplay = `${innerQty}`;
    } else {
      reviewFlag = 'REVIEW';
    }
  }

  /////////////////////////////////////
  // UNKNOWN
  /////////////////////////////////////
  else {
    reviewFlag = 'REVIEW';
  }

  let effectivePackPrice = '';
  let costPerUnit = '';
  let unitVal = toNumber(unitPrice);

  if (!isNaN(unitVal) && unitVal > 0) {
    effectivePackPrice = roundNumber(unitVal, 4);

    if (!isNaN(Number(packQty)) && Number(packQty) > 0) {
      costPerUnit = roundNumber(effectivePackPrice / Number(packQty), 6);
    }
  }

  if (!packQty || !baseUnit) {
    reviewFlag = 'REVIEW';
  }

  return {
    suggestedIngredient: cleaned,
    packSizeDisplay: packSizeDisplay,
    packQty: packQty,
    baseUnit: baseUnit,
    effectivePackPrice: effectivePackPrice,
    costPerUnit: costPerUnit,
    reviewFlag: reviewFlag,
    unitCount: unitCount
  };
}