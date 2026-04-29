/////////////////////////////////////
// MAKRO PARSER
/////////////////////////////////////
function parseMakroLine(rawDesc, packText, sizeText, packPrice, invoiceQty) {
  let text = (rawDesc || '').toString().trim();
  let packRaw = (packText || '').toString().trim();
  let sizeRaw = (sizeText || '').toString().trim();

  const upper = text.toUpperCase();

  // Skip VOID / CREDIT safely
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

  let outerQty = toNumber(packRaw);
  if (isNaN(outerQty) || outerQty <= 0) outerQty = 1;

  let m;

  function normaliseMakroCountUnit_(unitRaw) {
    const u = (unitRaw || '').toString().trim().toLowerCase();

    if (u === 'pk') return 'pk';
    if (u === 'ea' || u === 'unit' || u === 'units') return 'ea';
    if (u === 'roll' || u === 'rolls' || u === 'rol') return 'roll';
    if (u === 'can' || u === 'cans') return 'can';
    if (u === 'btl' || u === 'btls') return 'btl';
    if (u === 'sti' || u === 'stick' || u === 'sticks') return 'stick';
    if (u === 'sac' || u === 'sachet' || u === 'sachets') return 'sachet';
    if (u === 'ptn' || u === 'portion' || u === 'portions') return 'portion';

    return 'unit';
  }

  /////////////////////////////////////
  // WEIGHT / VOLUME
  // e.g. pack=24 size=330ml
  /////////////////////////////////////
  if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)$/i))) {
    const innerQty = Number(m[1]);
    const unitRaw = m[2];
    const unit = normaliseUnit(unitRaw);

    baseUnit = normaliseUnit(unit);
    if (baseUnit === 'kg') baseUnit = 'g';
    if (baseUnit === 'l') baseUnit = 'ml';

    packQty = roundNumber(convertToBaseQty(innerQty, unit) * outerQty, 4);
    packSizeDisplay = `${outerQty}x${m[1]}${unitRaw}`;
  }

  /////////////////////////////////////
  // COUNT ITEMS WITH UNIT
  // e.g. 1ea / 100pk / 6roll / 1000sti
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)\s*(pk|ea|unit|units|roll|rolls|rol|can|cans|btl|btls|sti|stick|sticks|sac|sachet|sachets|ptn|portion|portions)$/i))) {
    const innerQty = Number(m[1]);
    const unitLabel = normaliseMakroCountUnit_(m[2]);

    baseUnit = 'unit';
    packQty = roundNumber(innerQty * outerQty, 4);
    unitCount = packQty;
    packSizeDisplay = `${outerQty}x${innerQty}${unitLabel}`;
  }

  /////////////////////////////////////
  // COUNT ITEMS
  // e.g. 30s
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)\s*s$/i))) {
    const innerQty = Number(m[1]);

    baseUnit = 'unit';
    packQty = roundNumber(innerQty * outerQty, 4);
    unitCount = packQty;
    packSizeDisplay = `${outerQty}x${innerQty}s`;
  }

  /////////////////////////////////////
  // PLAIN COUNT
  // e.g. 1 / 10 / 100
  /////////////////////////////////////
  else if ((m = sizeRaw.match(/^(\d+(?:\.\d+)?)$/i))) {
    const innerQty = Number(m[1]);

    baseUnit = 'unit';
    packQty = roundNumber(innerQty * outerQty, 4);
    unitCount = packQty;
    packSizeDisplay = `${outerQty}x${innerQty}`;
  }

  else {
    reviewFlag = 'REVIEW';
  }

  let lineQty = toNumber(invoiceQty);
  if (!isNaN(lineQty) && lineQty > 1 && !isNaN(Number(packQty))) {
    packQty = roundNumber(Number(packQty) * lineQty, 4);
    if (baseUnit === 'unit') unitCount = packQty;
  }

  let effectivePackPrice = toNumber(packPrice);
  let costPerUnit = '';

  if (!isNaN(effectivePackPrice) && effectivePackPrice > 0 && !isNaN(Number(packQty)) && Number(packQty) > 0) {
    costPerUnit = roundNumber(effectivePackPrice / Number(packQty), 6);
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