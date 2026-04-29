/////////////////////////////////////
// FREESTONS PARSER
/////////////////////////////////////

function parseFreestonsLine(rawDesc, packPrice, invoiceQty) {
  const original = (rawDesc || '').toString().trim();
  const price = toNumber(packPrice);

  let cleaned = original;
  let packSizeDisplay = '';
  let packQty = '';
  let baseUnit = '';
  let reviewFlag = 'OK';
  let effectivePackPrice = price;

  let m;

  /////////////////////////////////////
  // 1) Read anything inside brackets
  // e.g. [12.5kg] / [2.5kg] / [40x190g] / [260]
  /////////////////////////////////////
  m = original.match(/\[([^\]]+)\]/);

  if (m) {
    const inside = (m[1] || '').toString().trim().toLowerCase();

    // A) Nested weight pack e.g. 40x190g
    let nested = inside.match(/^(\d+)\s*x\s*(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)$/i);
    if (nested) {
      const outer = Number(nested[1]);
      const inner = Number(nested[2]);
      const unit = normaliseUnit(nested[3]);

      packSizeDisplay = `${outer}x${stripTrailingZeros_(inner)}${unit}`;
      packQty = outer * convertToBaseQty(inner, unit);
      baseUnit = unitToBaseUnit(unit);
    }

    // B) Weight / volume only e.g. 12.5kg / 2.5kg / 5kg / 25kg
    if (!packQty) {
      let weight = inside.match(/^(\d+(?:\.\d+)?)\s*(kg|g|ml|l|ltr|litre|liter)$/i);
      if (weight) {
        const qty = Number(weight[1]);
        const unit = normaliseUnit(weight[2]);

        packSizeDisplay = `${stripTrailingZeros_(qty)}${unit}`;
        packQty = convertToBaseQty(qty, unit);
        baseUnit = unitToBaseUnit(unit);
      }
    }

    // C) Count only e.g. [260]
    if (!packQty) {
      let count = inside.match(/^(\d+)$/);
      if (count) {
        const qty = Number(count[1]);

        packSizeDisplay = `${qty}ea`;
        packQty = qty;
        baseUnit = 'unit';
      }
    }
  }

  /////////////////////////////////////
  // 2) Clean ingredient name hard
  /////////////////////////////////////
  cleaned = original
    .replace(/\[[^\]]*\]/g, ' ')   // remove all [ ... ]
    .replace(/\[\]/g, ' ')         // remove empty []
    .replace(/\s+/g, ' ')
    .trim();

  let ingredient = toTitleCase(cleaned)
    .replace(/\[[^\]]*\]/g, ' ')
    .replace(/\[\]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const cleanName = ingredient.toLowerCase();

  /////////////////////////////////////
  // 3) Cost per unit
  /////////////////////////////////////
  if (isNaN(Number(effectivePackPrice))) {
  reviewFlag = 'CHECK PRICE';
    }

    let costPerUnit = '';
    if (!isNaN(effectivePackPrice) && Number(packQty) > 0) {
      costPerUnit = roundNumber(effectivePackPrice / Number(packQty), 6);
    }

  /////////////////////////////////////
  // 4) Review flag
  /////////////////////////////////////
  if (!ingredient || !packQty || !baseUnit || isNaN(effectivePackPrice)) {
    reviewFlag = 'CHECK';
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