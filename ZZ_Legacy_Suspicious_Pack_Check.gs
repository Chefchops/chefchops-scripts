//////////////////////////////////////////
// SUSPICIOUS PACK SIZE CHECK
//////////////////////////////////////////
function isSuspiciousPackSize_(ingredientName, packQty, baseUnit, packSizeText) {
  const name = (ingredientName || '').toString().trim().toLowerCase();
  const unit = (baseUnit || '').toString().trim().toLowerCase();
  const sizeText = (packSizeText || '').toString().trim().toLowerCase();
  const qty = toNumber(packQty);

  if (isNaN(qty) || qty <= 0) return false;

  /////////////////////////////////////////
  // ALLOWED EXCEPTIONS
  /////////////////////////////////////////
  const isOil =
    name.indexOf('oil') !== -1 ||
    sizeText.indexOf('oil') !== -1;

  const isPotato =
    name.indexOf('potato') !== -1 ||
    name.indexOf('potatoes') !== -1;

  // Allow big vegetable oil packs
  if (isOil && unit === 'ml' && qty <= 20000) return false;
  if (isOil && unit === 'l' && qty <= 20) return false;

  // Allow big potato sacks
  if (isPotato && unit === 'g' && qty <= 25000) return false;
  if (isPotato && unit === 'kg' && qty <= 25) return false;

  /////////////////////////////////////////
  // GENERAL SAFETY RULES
  /////////////////////////////////////////
  // Anything over 10kg / 10ltr should be reviewed
  if (unit === 'g' && qty > 10000) return true;
  if (unit === 'kg' && qty > 10) return true;

  if (unit === 'ml' && qty > 10000) return true;
  if (unit === 'l' && qty > 10) return true;

  return false;
}