// ✅ Global Constants - Folder IDs
const FOLDER_ID_RESI = "1XnMUkIPdzJpDZGL_F3wOw2Tu7JaJUuj-"; // Folder Resi PDF
const FOLDER_ID_TRANSFER = "1Eg0uhsoDXdcOB0VirOBDD9MmVc6KHhjz"; // Folder Bukti Transfer


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var editedCell = e.range;

  if (editedCell.getColumn() == 4) { // Kolom D, indeks kolom mulai dari 1
    var rangeE = sheet.getRange(editedCell.getRow(), 5); // Kolom E
    var rangeH = sheet.getRange(editedCell.getRow(), 8); // Kolom H

    if (editedCell.getValue() == 'RTB') {
      rangeE.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('RTB'))
        .build());
      rangeH.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('ListLokasi!A1:A2')) // Jakarta, Yogya
        .build());
    } else if (editedCell.getValue() == 'PTB') {
      rangeE.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('PTB'))
        .build());
      rangeH.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRangeByName('ListLokasi!A3')) // Online
        .build());
    } else {
      rangeE.clearDataValidations();
      rangeH.clearDataValidations();
    }
  }
}