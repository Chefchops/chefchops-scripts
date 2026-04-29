//////////////////////////////////////
// PARSER DISPATCHER
/////////////////////////////////////
function parseInvoiceLineBySupplier(
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
) {
  const supplierKey = (supplier || '').toString().trim().toLowerCase();

  if (supplierKey.includes('freeston')) {
    return parseFreestonsLine(rawDesc, price, invoiceQty);
  }

  if (supplierKey.indexOf('bidfood') !== -1) {
    return parseBidfoodLine(rawDesc, packSize, price, invoiceQty, qtyType, unitsWeight);
  }

  if (supplierKey.indexOf('easter') !== -1) {
    return parseEastersLine(rawDesc, price, invoiceQty, qtyType);
  }

  if (supplierKey.indexOf('hazel') !== -1) {
    return parseHazelsLine(rawDesc, price, invoiceQty, qtyType);
  }

  if (supplierKey.indexOf('broadland') !== -1) {
    return parseBroadlandLine(rawDesc, price, invoiceQty, qtyType);
  }

  if (supplierKey.indexOf('makro') !== -1) {
    return parseMakroLine(rawDesc, packSize, sizeText, price, invoiceQty);
  }

  if (supplierKey.indexOf('pilgrim') !== -1) {
    return parsePilgrimLine(rawDesc, caseSize, packSize, price, casesOrdered, singlesOrdered);
  }

  return parseInvoiceLine(rawDesc, price, invoiceQty, qtyType);
}