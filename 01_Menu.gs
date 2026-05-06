/////////////////////////////////////
// ON OPEN
/////////////////////////////////////

function onOpen() {
  buildChefchopsMenu_();
}


/////////////////////////////////////
// CHEFCHOPS MENU
// CLEAN CURRENT LIVE PIPELINE ONLY
/////////////////////////////////////

function buildChefchopsMenu_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Chefchops')

    /////////////////////////////////////
    // PDF PIPELINE
    /////////////////////////////////////

    .addSubMenu(
      ui.createMenu('PDF Pipeline')

        .addItem('1. Import PDFs From Drive', 'importPdfJobsFromDriveFolder')
        .addItem('2. Process Last PDF (Cloud)', 'processLastPdfRow_')

        .addSeparator()

        .addItem('3. Build Header + Extracted Lines + Review', 'runBuildExtractedLinesFromPdfJson')

        .addSeparator()

        .addItem('4. Append Reviewed PDF to Ingredients Master', 'appendReviewedPdfExtractedLinesToIngredientsMaster')
    )

    /////////////////////////////////////
    // PDF REVIEW
    /////////////////////////////////////

    .addSubMenu(
      ui.createMenu('PDF Review')

        .addItem('Setup Review Sheet', 'setupPdfReviewSheet')
        .addItem('Apply Review Corrections', 'applyPdfReviewCorrections')
        .addItem('Highlight Missing Fields', 'highlightPdfReviewMissingFields')

        .addSeparator()

        .addItem('Clear Review Sheet', 'clearPdfReviewSheet')
    )
/////////////////////////////////////
// PRICE TOOLS
/////////////////////////////////////

    .addSubMenu(
      ui.createMenu('Price Tools')
        .addItem('Setup Price Search', 'setupPriceSearchSheet')
        .addItem('Run Price Search', 'runPriceSearch')
        .addItem('Clear Price Search', 'clearPriceSearchResults')
    )
    .addToUi();
}