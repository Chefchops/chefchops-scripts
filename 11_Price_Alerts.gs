/////////////////////////////////////
// PRICE ALERT HELPERS
/////////////////////////////////////
function logPriceChange_(oldPrice, newPrice, ingredient, supplier, site, alerts) {
  const oldVal = toNumber(oldPrice);
  const newVal = toNumber(newPrice);

  if (isNaN(oldVal) || isNaN(newVal)) return;
  if (oldVal === newVal) return;

  const diff = roundNumber(newVal - oldVal, 4);
  const direction = diff > 0 ? 'Increase' : 'Decrease';

  alerts.push([
    new Date(),                 // A Date
    site || '',                 // B Site
    supplier || '',             // C Supplier
    ingredient || '',           // D Ingredient
    oldVal,                     // E Old Price
    newVal,                     // F New Price
    diff,                       // G Difference
    direction                   // H Change Type
  ]);
}

function writePriceAlertsSheet_(alerts) {
  const ss = SpreadsheetApp.getActive();
  const sheetName = 'Price Alerts';

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clearContents();
  sheet.clearFormats();

  const headers = [[
    'Date',
    'Site',
    'Supplier',
    'Ingredient',
    'Old Price (£)',
    'New Price (£)',
    'Change (£)',
    'Direction'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold');

  if (alerts.length > 0) {
    sheet.getRange(2, 1, alerts.length, headers[0].length).setValues(alerts);

    sheet.getRange(2, 1, alerts.length, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange(2, 5, alerts.length, 3).setNumberFormat('£0.00');

    const values = sheet.getRange(2, 8, alerts.length, 1).getValues();

    for (let i = 0; i < values.length; i++) {
      const dir = values[i][0];
      if (dir === 'Increase') {
        sheet.getRange(i + 2, 1, 1, 8).setBackground('#fce8e6');
      } else if (dir === 'Decrease') {
        sheet.getRange(i + 2, 1, 1, 8).setBackground('#e6f4ea');
      }
    }
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 8);
}