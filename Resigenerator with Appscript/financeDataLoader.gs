// ✅ Fungsi bantu untuk merapikan nama file (hilangkan spasi berlebih)
function normalizeWhitespace(str) {
  return str.toString().replace(/\s+/g, ' ').trim();
}

// ✅ Fungsi utama muat ulang data ke sheet "Data Email ke Finance"
function loadFinanceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Form Responses 1");
  const targetSheet = ss.getSheetByName("Data Email ke Finance");

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert("❌ Sheet tidak ditemukan.");
    return;
  }

  // ✅ Folder Transfer dari global Code.gs
  const folder = DriveApp.getFolderById(FOLDER_ID_TRANSFER);
  const sourceData = sourceSheet.getDataRange().getValues();
  const headerRow = sourceData[0];

  // ✅ Fungsi cari kolom by nama
  function findColumnIndex(headerRow, keyword) {
    return headerRow.findIndex(h => String(h).toLowerCase().includes(keyword.toLowerCase()));
  }

  const IDX_TRX       = findColumnIndex(headerRow, "ID Transaksi");
  const IDX_NAMA      = findColumnIndex(headerRow, "Nama Lengkap");
  const IDX_TIPE      = findColumnIndex(headerRow, "Tipe Program");
  const IDX_PROGRAM   = findColumnIndex(headerRow, "Nama Program");
  const IDX_TANGGAL   = findColumnIndex(headerRow, "Otomatis Jangan");
  const IDX_JUMLAH    = findColumnIndex(headerRow, "Jumlah Pembayaran");
  const IDX_STATUS    = findColumnIndex(headerRow, "Status Pembayaran");
  const IDX_CHANNEL   = findColumnIndex(headerRow, "Channel Pembayaran");
  const IDX_EMAIL     = findColumnIndex(headerRow, "Send Email Status");
  const IDX_FILE_RESI = findColumnIndex(headerRow, "File dalam Folder");

  const requiredIndexes = [
    IDX_TRX, IDX_NAMA, IDX_TIPE, IDX_PROGRAM,
    IDX_TANGGAL, IDX_JUMLAH, IDX_STATUS, IDX_CHANNEL,
    IDX_EMAIL, IDX_FILE_RESI
  ];
  if (requiredIndexes.includes(-1)) {
    SpreadsheetApp.getUi().alert("❌ Kolom tidak lengkap.");
    return;
  }

  const outputHeaders = [
    "No",
    "ID Transaksi",
    "Nama Lengkap",
    "Tipe Program",
    "Nama Program",
    "Tanggal Transaksi",
    "Jumlah Pembayaran",
    "Status Pembayaran",
    "Channel Pembayaran",
    "Status Bukti Transfer",
    "Status Kirim ke Finance"
  ];

  // 🛡️ Proteksi baris header A1:K1 hanya jika belum ada
  try {
    const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const hasHeaderProtection = protections.some(p => p.getRange().getA1Notation() === "A1:K1");
    if (!hasHeaderProtection) {
      const protection = targetSheet.getRange("A1:K1").protect().setDescription("Header Lock");
      protection.setWarningOnly(true); // ⚠️ Aman untuk Editor
    }
  } catch (e) {
    Logger.log("⚠️ Proteksi header dilewati: " + e.message);
  }

  // ⛑ Backup "Status Kirim ke Finance" dari data lama
  const existingRows = targetSheet.getLastRow() - 1;
  const oldStatusData = existingRows > 0
    ? targetSheet.getRange(2, 11, existingRows, 1).getValues()
    : [];

  const outputData = [];
  let no = existingRows + 1;
  let oldStatusIndex = 0;

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const emailStatus = String(row[IDX_EMAIL] || "").trim();
    const fileResiStatus = String(row[IDX_FILE_RESI] || "").trim();

    if (emailStatus.startsWith("Sending completed ✅") && fileResiStatus === "Ada") {
      const idTransaksi = row[IDX_TRX];
      const nama = row[IDX_NAMA];
      const tipe = row[IDX_TIPE];
      const program = row[IDX_PROGRAM];
      const status = row[IDX_STATUS];
      const clean = normalizeWhitespace;

      const fileName = `[Phincon Academy] Bukti Transfer - ${clean(idTransaksi)} - ${clean(nama)} - ${clean(tipe)} ${clean(program)} - ${clean(status)}`;

      let statusBukti = "Not yet uploaded";
      if (folder.getFilesByName(fileName).hasNext()) {
        statusBukti = "✅ Uploaded";
      }

      const statusKirim = oldStatusData[oldStatusIndex] ? oldStatusData[oldStatusIndex][0] : "";
      oldStatusIndex++;

      outputData.push([
        no++,
        idTransaksi,
        nama,
        tipe,
        program,
        row[IDX_TANGGAL],
        row[IDX_JUMLAH],
        status,
        row[IDX_CHANNEL],
        statusBukti,
        statusKirim || ""
      ]);
    }
  }

  // 🧾 Tulis header jika belum ada sama sekali
  if (targetSheet.getLastRow() === 0) {
    targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  }

  // ✅ Tambahkan baris baru tanpa mengganggu data sebelumnya
  if (outputData.length > 0) {
    const startRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
  }

  SpreadsheetApp.getUi().alert("✅ Data berhasil dimuat dan siap dikirim ke Finance");
}



// // ✅ Fungsi bantu untuk merapikan nama file (hilangkan spasi berlebih)
// function normalizeWhitespace(str) {
//   return str.toString().replace(/\s+/g, ' ').trim();
// }

// // ✅ Fungsi utama muat ulang data ke sheet "Data Email ke Finance"
// function loadFinanceData() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sourceSheet = ss.getSheetByName("Form Responses 1");
//   const targetSheet = ss.getSheetByName("Data Email ke Finance");

//   if (!sourceSheet || !targetSheet) {
//     SpreadsheetApp.getUi().alert("❌ Sheet tidak ditemukan.");
//     return;
//   }

//   // ✅ Folder Transfer dari global Code.gs
//   const folder = DriveApp.getFolderById(FOLDER_ID_TRANSFER);
//   const sourceData = sourceSheet.getDataRange().getValues();
//   const headerRow = sourceData[0];

//   // ✅ Fungsi cari kolom by nama
//   function findColumnIndex(headerRow, keyword) {
//     return headerRow.findIndex(h => String(h).toLowerCase().includes(keyword.toLowerCase()));
//   }

//   const IDX_TRX       = findColumnIndex(headerRow, "ID Transaksi");
//   const IDX_NAMA      = findColumnIndex(headerRow, "Nama Lengkap");
//   const IDX_TIPE      = findColumnIndex(headerRow, "Tipe Program");
//   const IDX_PROGRAM   = findColumnIndex(headerRow, "Nama Program");
//   const IDX_TANGGAL   = findColumnIndex(headerRow, "Otomatis Jangan");
//   const IDX_JUMLAH    = findColumnIndex(headerRow, "Jumlah Pembayaran");
//   const IDX_STATUS    = findColumnIndex(headerRow, "Status Pembayaran");
//   const IDX_CHANNEL   = findColumnIndex(headerRow, "Channel Pembayaran");
//   const IDX_EMAIL     = findColumnIndex(headerRow, "Send Email Status");
//   const IDX_FILE_RESI = findColumnIndex(headerRow, "File dalam Folder");

//   const requiredIndexes = [
//     IDX_TRX, IDX_NAMA, IDX_TIPE, IDX_PROGRAM,
//     IDX_TANGGAL, IDX_JUMLAH, IDX_STATUS, IDX_CHANNEL,
//     IDX_EMAIL, IDX_FILE_RESI
//   ];
//   if (requiredIndexes.includes(-1)) {
//     SpreadsheetApp.getUi().alert("❌ Kolom tidak lengkap.");
//     return;
//   }

//   const outputHeaders = [
//     "No",
//     "ID Transaksi",
//     "Nama Lengkap",
//     "Tipe Program",
//     "Nama Program",
//     "Tanggal Transaksi",
//     "Jumlah Pembayaran",
//     "Status Pembayaran",
//     "Channel Pembayaran",
//     "Status Bukti Transfer",
//     "Status Kirim ke Finance"
//   ];

//   // 🛡️ Proteksi baris header A1:K1 hanya jika belum ada
//   try {
//     const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//     const hasHeaderProtection = protections.some(p => p.getRange().getA1Notation() === "A1:K1");
//     if (!hasHeaderProtection) {
//       const protection = targetSheet.getRange("A1:K1").protect().setDescription("Header Lock");
//       protection.setWarningOnly(true); // ⚠️ Aman untuk Editor
//     }
//   } catch (e) {
//     Logger.log("⚠️ Proteksi header dilewati: " + e.message);
//   }

//   // ⛑ Backup "Status Kirim ke Finance" dari data lama
//   const existingRows = targetSheet.getLastRow() - 1;
//   const oldStatusData = existingRows > 0
//     ? targetSheet.getRange(2, 11, existingRows, 1).getValues()
//     : [];

//   const outputData = [];
//   let no = existingRows + 1;
//   let oldStatusIndex = 0;

//   for (let i = 1; i < sourceData.length; i++) {
//     const row = sourceData[i];
//     const emailStatus = String(row[IDX_EMAIL] || "").trim();
//     const fileResiStatus = String(row[IDX_FILE_RESI] || "").trim();

//     if (emailStatus.startsWith("Sending completed ✅") && fileResiStatus === "Ada") {
//       const idTransaksi = row[IDX_TRX];
//       const nama = row[IDX_NAMA];
//       const tipe = row[IDX_TIPE];
//       const program = row[IDX_PROGRAM];
//       const status = row[IDX_STATUS];
//       const clean = normalizeWhitespace;

//       const fileName = `[Phincon Academy] Bukti Transfer - ${clean(idTransaksi)} - ${clean(nama)} - ${clean(tipe)} ${clean(program)} - ${clean(status)}`;

//       let statusBukti = "Not yet uploaded";
//       if (folder.getFilesByName(fileName).hasNext()) {
//         statusBukti = "✅ Uploaded";
//       }

//       const statusKirim = oldStatusData[oldStatusIndex] ? oldStatusData[oldStatusIndex][0] : "";
//       oldStatusIndex++;

//       outputData.push([
//         no++,
//         idTransaksi,
//         nama,
//         tipe,
//         program,
//         row[IDX_TANGGAL],
//         row[IDX_JUMLAH],
//         status,
//         row[IDX_CHANNEL],
//         statusBukti,
//         statusKirim || ""
//       ]);
//     }
//   }

//   // 🧾 Tulis header jika belum ada sama sekali
//   if (targetSheet.getLastRow() === 0) {
//     targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
//   }

//   // ✅ Tambahkan baris baru tanpa mengganggu data sebelumnya
//   if (outputData.length > 0) {
//     const startRow = targetSheet.getLastRow() + 1;
//     targetSheet.getRange(startRow, 1, outputData.length, outputHeaders.length).setValues(outputData);
//   }

//   SpreadsheetApp.getUi().alert("✅ Data berhasil dimuat dan siap dikirim ke Finance");
// }