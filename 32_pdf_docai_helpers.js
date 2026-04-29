/////////////////////////////////////
// DOC AI HELPERS
/////////////////////////////////////

function getTextFromLayout_(doc, layout) {
  if (!layout || !layout.textAnchor || !layout.textAnchor.textSegments) return '';

  return layout.textAnchor.textSegments
    .map(seg => doc.text.substring(seg.startIndex || 0, seg.endIndex))
    .join('');
}

function extractTablesFromDocAI_(doc) {
  const tables = [];

  (doc.pages || []).forEach(page => {
    (page.tables || []).forEach(table => {
      const rows = [];

      const allRows = []
        .concat(table.headerRows || [])
        .concat(table.bodyRows || []);

      allRows.forEach(row => {
        const cells = (row.cells || []).map(cell => getTextFromLayout_(doc, cell.layout));
        rows.push(cells);
      });

      tables.push(rows);
    });
  });

  return tables;
}
