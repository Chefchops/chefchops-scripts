/////////////////////////////////////
// APPLY PRICE COLOUR FORMATTING
// Green = cheapest
// Amber = slightly higher
// Red = higher
/////////////////////////////////////

function applyPriceColouring() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Ingredients Master');
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Missing "Ingredients Master" sheet.');
    return;
  }

  const lastRow = getLastUsedRowInColumn(sheet, 2, 2); // based on Clean Name
  if (lastRow < 2) {
    ui.alert('No data to format.');
    return;
  }

  const rowCount = lastRow - 1;

  // H = Cost per Unit (£)
  const range = sheet.getRange(2, 8, rowCount, 1);

  // Clear existing rules first
  const rules = sheet.getConditionalFormatRules().filter(r => {
    return !r.getRanges().some(rg => rg.getA1Notation() === range.getA1Notation());
  });

  /////////////////////////////////////
  // GREEN (cheapest = P = 0)
  /////////////////////////////////////
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$P2=0')
      .setBackground('#c6efce')
      .setFontColor('#006100')
      .setRanges([range])
      .build()
  );

  /////////////////////////////////////
  // AMBER (within small margin)
  /////////////////////////////////////
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($P2>0,$P2<=0.01)')
      .setBackground('#fff2cc')
      .setFontColor('#7f6000')
      .setRanges([range])
      .build()
  );

  /////////////////////////////////////
  // RED (more expensive)
  /////////////////////////////////////
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$P2>0.01')
      .setBackground('#ffc7ce')
      .setFontColor('#9c0006')
      .setRanges([range])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  ui.alert('Price colouring applied.');
}
