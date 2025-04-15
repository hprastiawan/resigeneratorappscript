// ✅ Fungsi untuk membuka sidebar upload bukti transfer dari baris aktif
function openUploadDialogForActiveRow() {
  const html = HtmlService.createHtmlOutputFromFile("uploadBuktiTransferWeb")
    .setWidth(400)
    .setHeight(300)
    .setTitle("Upload Bukti Transfer");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ✅ Ambil data baris aktif dari sheet "Kirim ke Finance"
function getActiveRowData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kirim ke Finance");
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return null;

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    rowIndex: row,
    idTransaksi: data[1],
    nama: data[2],
    tipe: data[3],
    program: data[4],
    status: data[7], // status pembayaran
  };
}

// ✅ Upload bukti transfer ke Google Drive berdasarkan data baris aktif
function uploadBuktiTransferFromDialog(blob, rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Kirim ke Finance");
  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  const clean = txt => String(txt || "").trim().replace(/\s+/g, " ");
  const fileName = `[Phincon Academy] Bukti Transfer - ${clean(row[1])} - ${clean(row[2])} - ${clean(row[3])} ${clean(row[4])} - ${clean(row[7])}`;

  const folder = DriveApp.getFolderById("1Eg0uhsoDXdcOB0VirOBDD9MmVc6KHhjz"); // ✅ Folder Bukti Transfer
  const file = folder.createFile(blob.setName(fileName));

  // ✅ Update kolom "Status Bukti Transfer" (kolom ke-10)
  sheet.getRange(rowIndex, 10).setValue("✅ Uploaded");

  // 🔔 Tampilkan alert sukses
  SpreadsheetApp.getUi().alert("📤 Upload Bukti Bayar", `1 file berhasil diupload:\n${fileName}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ✅ Decode base64 lalu teruskan ke fungsi upload utama
function uploadBase64File(base64Data, fileName, mimeType, rowIndex) {
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  return uploadBuktiTransferFromDialog(blob, rowIndex);
}



// ====================================
// KE SHEET Data Email ke Finance
// ====================================

// function openUploadDialogForActiveRow() {
//   const html = HtmlService.createHtmlOutputFromFile("uploadBuktiTransferWeb")
//     .setWidth(400)
//     .setHeight(300)
//     .setTitle("Upload Bukti Transfer");
//   SpreadsheetApp.getUi().showSidebar(html);
// }

// function getActiveRowData() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Email ke Finance");
//   const row = sheet.getActiveCell().getRow();
//   if (row < 2) return null;

//   const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
//   return {
//     rowIndex: row,
//     idTransaksi: data[1],
//     nama: data[2],
//     tipe: data[3],
//     program: data[4],
//     status: data[7],
//   };
// }

// function uploadBuktiTransferFromDialog(blob, rowIndex) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("Data Email ke Finance");
//   const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

//   const clean = txt => String(txt || "").trim().replace(/\s+/g, " ");
//   const fileName = `[Phincon Academy] Bukti Transfer - ${clean(row[1])} - ${clean(row[2])} - ${clean(row[3])} ${clean(row[4])} - ${clean(row[7])}`;

//   const folder = DriveApp.getFolderById("1Eg0uhsoDXdcOB0VirOBDD9MmVc6KHhjz");
//   const file = folder.createFile(blob.setName(fileName));

//   // Update status kolom "Status Bukti Transfer" (kolom ke-10)
//   sheet.getRange(rowIndex, 10).setValue("✅ Uploaded");

//   // 🔔 Tampilkan alert box di tengah layar
//   SpreadsheetApp.getUi().alert("📤 Upload Bukti Bayar", `1 file berhasil diupload:\n${fileName}`, SpreadsheetApp.getUi().ButtonSet.OK);
// }


// function uploadBase64File(base64Data, fileName, mimeType, rowIndex) {
//   const decoded = Utilities.base64Decode(base64Data);
//   const blob = Utilities.newBlob(decoded, mimeType, fileName);
//   return uploadBuktiTransferFromDialog(blob, rowIndex);
// }

