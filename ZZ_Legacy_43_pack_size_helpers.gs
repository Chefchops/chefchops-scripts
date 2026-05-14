/////////////////////////////////////
// PACK SIZE HELPERS
// LEGACY FUNCTION NAME
// NOW ROUTES TO PACK SIZE STANDARD
/////////////////////////////////////

function parsePackSizeToUnits_(packSize) {
  const parsed = parsePackSizeToUnitsStandard_(packSize);

  return {
    packQty: parsed.packQty,
    baseUnit: parsed.baseUnit,
    unitPerCase: parsed.unitPerCase,
    unitPerPackCase: parsed.unitPerPackCase,
    reviewFlag: parsed.reviewFlag,
    notes: parsed.notes,
    displayPackSize: parsed.displayPackSize,
    cleanedPackSize: parsed.cleanedPackSize
  };
}