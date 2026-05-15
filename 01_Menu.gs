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
      ui
        .createMenu('PDF Pipeline')

        /////////////////////////////////////
        // STEP 1 - IMPORT
        /////////////////////////////////////

        .addItem('1. Import PDFs From Drive', 'importPdfJobsFromDriveFolder')

        .addSeparator()

        /////////////////////////////////////
        // STEP 2 - CLOUD PROCESSING
        /////////////////////////////////////

        .addItem('2. Process Next 5 PDFs (Cloud)', 'processNext5PendingPdfRows')

        .addSeparator()

        /////////////////////////////////////
        // STEP 3 - BUILD HEADER + REVIEW
        /////////////////////////////////////

        .addItem(
          '3. Build Next 5 Stored PDFs + Review',
          'buildNext5StoredPdfJsonFiles',
        )

        .addItem(
          '3.1 Build Latest Stored PDF + Review',
          'runBuildExtractedLinesFromPdfJson',
        )

        .addSeparator()

        /////////////////////////////////////
        // STEP 4 - CORRECTIONS + APPEND
        /////////////////////////////////////

        .addItem(
          '4. Apply Corrections + Append to Ingredients Master',
          'applyReviewCorrectionsThenAppendPdf',
        ),
    )

    /////////////////////////////////////
    // PDF REVIEW
    /////////////////////////////////////

    .addSubMenu(
      ui
        .createMenu('PDF Review')

        .addItem('Setup Review Sheet', 'setupPdfReviewSheet')
        .addItem('Apply Review Corrections Only', 'applyPdfReviewCorrections')
        .addItem('Highlight Missing Fields', 'highlightPdfReviewMissingFields')

        .addSeparator()

        .addItem('Clear Review Sheet', 'clearPdfReviewSheet'),
    )

    /////////////////////////////////////
    // PRICE TOOLS
    /////////////////////////////////////

    .addSubMenu(
      ui
        .createMenu('Price Tools')

        .addItem('Setup Price Search', 'setupPriceSearchSheet')
        .addItem('Run Price Search', 'runPriceSearch')
        .addItem('Clear Price Search', 'clearPriceSearchResults'),
    )

    .addToUi();
}
