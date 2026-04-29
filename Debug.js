function testFirstBidfoodRowKeys() {
  const fileId = Browser.inputBox('Enter Drive File ID');

  if (!fileId || fileId === 'cancel') return;

  const json = rebuildJsonFromChunks_(fileId);
  const row = (json.bidfoodRows || [])[0];

  Logger.log(JSON.stringify(row, null, 2));

  SpreadsheetApp.getUi().alert(
    'First row keys:\n\n' + Object.keys(row).join('\n')
  );
}

function testOpenPdfFolder() {
  const folderId = '1uSoRK8QrqjRSXoBBrJWDDn2S576Js-Mv';

  try {
    const folder = DriveApp.getFolderById(folderId);
    SpreadsheetApp.getUi().alert(
      'Folder opened OK:\n\n' + folder.getName()
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Could not open folder:\n\n' + err.message
    );
  }
}




