/////////////////////////////////////
// ON OPEN
/////////////////////////////////////

function onOpen() {
  buildChefchopsMenu_();
}

/////////////////////////////////////
// CHEFCHOPS MENU (CLEAN + ORDERED)
/////////////////////////////////////

function buildChefchopsMenu_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Chefchops')

    /////////////////////////////////////
    // PDF PIPELINE (NEW)
    /////////////////////////////////////
    .addSubMenu(
      ui.createMenu('PDF Pipeline')

        .addItem('1. Import PDFs From Drive', 'importPdfJobsFromDriveFolder')
        .addItem('2. Process Last PDF (Cloud)', 'processLastPdfRow_')

        .addSeparator()

        .addItem('3. Build Extracted Lines + Review', 'runBuildExtractedLinesFromPdfJson')

        .addSeparator()

        .addItem('4. Build Invoice Import From PDF Review', 'buildInvoiceImportFromPdfReview')
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
    // INGREDIENTS
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
    // COMPARISON
    /////////////////////////////////////
    .addSubMenu(
      ui.createMenu('Comparison')
        .addItem('Build Supplier Comparison', 'buildSupplierComparison')
    )

    .addToUi();
}