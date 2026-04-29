/////////////////////////////////////
// BROADLAND PARSER
/////////////////////////////////////
function parseBroadlandLine(rawDesc, packPrice, invoiceQty, qtyType) {
  const original = (rawDesc || '').toString().trim();
  let text = original;

  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let reviewFlag = 'OK';
  let effectivePackPrice = toNumber(packPrice);
  let suggestedIngredient = '';
  let cleanName = '';
  let matchedRule = false;

  const broadlandRules = [
    {
      match: /sweetbriar.*unsmoked.*sliced back/i,
      ingredient: 'Sweetbriar Unsmoked Sliced Back Bacon',
      packSizeDisplay: '2.27kg',
      packQty: 2270,
      baseUnit: 'g'
    }
  ];

  for (let i = 0; i < broadlandRules.length; i++) {
    const rule = broadlandRules[i];
    if (rule.match && rule.match.test(text)) {
      matchedRule = true;
      suggestedIngredient = rule.ingredient || '';
      cleanName = suggestedIngredient ? suggestedIngredient.toLowerCase().trim() : '';
      packSizeDisplay = rule.packSizeDisplay || '';
      packQty = rule.packQty || '';
      baseUnit = rule.baseUnit || '';

      effectivePackPrice = toNumber(packPrice);
      reviewFlag = 'OK';
      break;
    }
  }

  if (!matchedRule) {
    let cleaned = text
      .replace(/\s+/g, ' ')
      .trim();

    cleaned = cleaned
      .replace(/\bBOX\b\s*/i, '')
      .replace(/\bPACK\b\s*/i, '')
      .trim();

    let m = cleaned.match(/\b(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)\b/i);

    if (m) {
      const qty = Number(m[1]);
      const unit = normaliseUnit(m[2]);

      packSizeDisplay = `${stripTrailingZeros_(qty)}${unit}`;
      baseUnit = unitToBaseUnit(unit);
      packQty = convertToBaseQty(qty, unit);

      cleaned = cleaned.replace(m[0], '').replace(/\s+/g, ' ').trim();
      suggestedIngredient = toTitleCase(cleaned);
      cleanName = suggestedIngredient.toLowerCase().trim();

      effectivePackPrice = toNumber(packPrice);
    } else {
      suggestedIngredient = toTitleCase(cleaned);
      cleanName = suggestedIngredient.toLowerCase().trim();

      reviewFlag = 'CHECK UNIT';
      packSizeDisplay = qtyType ? qtyType.toString().trim().toLowerCase() : 'box';
      packQty = '';
      baseUnit = '';
    }
  }

  if (isNaN(Number(effectivePackPrice))) {
  reviewFlag = 'CHECK PRICE';
    }

    let costPerUnit = '';
    if (effectivePackPrice && packQty) {
      costPerUnit = roundNumber(effectivePackPrice / packQty, 6);
    }

  if (!suggestedIngredient) {
    reviewFlag = 'CHECK';
  }

  if (!packQty || !baseUnit) {
    if (reviewFlag === 'OK') reviewFlag = 'CHECK UNIT';
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