/////////////////////////////////////
// ON OPEN
/////////////////////////////////////

function onOpen() {
  buildChefchopsMenu_();
}

/////////////////////////////////////
// CHEFCHOPS MENU
// CLEAN CURRENT PIPELINE MENU
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

        .addItem('3. Build Extracted Lines + Review', 'runBuildExtractedLinesFromPdfJson')

        .addSeparator()

        .addItem('4. Append Reviewed PDF to Ingredients Master', 'appendReviewedPdfExtractedLinesToIngredientsMaster')
    )

    /////////////////////////////////////
    // PDF HEADERS
    /////////////////////////////////////
    .addSubMenu(
      ui.createMenu('PDF Headers')

        .addItem('Setup Headers Sheet', 'setupPdfInvoiceHeadersSheet')
        .addItem('Build Latest Invoice Header', 'buildLatestPdfInvoiceHeader')
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
    // INGREDIENTS SUPPORT
    /////////////////////////////////////
    .addSubMenu(
      ui.createMenu('Ingredients')

        .addItem('Check Missing Categories', 'checkMissingCategories')
        .addItem('Check Missing Product Groups', 'checkMissingProductGroups')
        .addItem('Auto Suggest Product Groups', 'autoSuggestIngredientProductGroups')

        .addSeparator()

        .addItem('Set Category Dropdown', 'setIngredientCategoryDropdown')
        .addItem('Set Product Group Dropdown', 'setIngredientProductGroupDropdown')
    )

    /////////////////////////////////////
    // SUPPLIER COMPARISON
    /////////////////////////////////////
    .addSubMenu(
      ui.createMenu('Comparison')
        .addItem('Build Supplier Comparison', 'buildSupplierComparison')
    )

    .addToUi();
}