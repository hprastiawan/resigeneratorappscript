// ✅ Fungsi bantu hilangkan spasi berlebih di nama file
function normalizeWhitespace(str) {
  return str.toString().replace(/\s+/g, ' ').trim();
}

// ✅ Fungsi utama muat ulang data ke sheet "Kirim ke Finance"
function loadDataKirimKeFinance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Form Responses 1");
  const targetSheet = ss.getSheetByName("Kirim ke Finance");

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert("❌ Sheet tidak ditemukan.");
    return;
  }

  const folder = DriveApp.getFolderById(FOLDER_ID_TRANSFER);
  const sourceData = sourceSheet.getDataRange().getValues();
  const headerRow = sourceData[0];

  function findColumnIndex(header, keyword) {
    return header.findIndex(h => String(h).toLowerCase().includes(keyword.toLowerCase()));
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
    IDX_TRX, IDX_NAMA, IDX_TIPE, IDX_PROGRAM, IDX_TANGGAL,
    IDX_JUMLAH, IDX_STATUS, IDX_CHANNEL, IDX_EMAIL, IDX_FILE_RESI
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

  // 🛡️ Proteksi baris header hanya jika belum ada (Editor tetap bisa jalan)
  try {
    const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    const hasHeader = protections.some(p => p.getRange().getA1Notation() === "A1:K1");
    if (!hasHeader) {
      const protect = targetSheet.getRange("A1:K1").protect().setDescription("Header Lock");
      protect.setWarningOnly(true);
    }
  } catch (err) {
    Logger.log("⚠️ Proteksi header dilewati: " + err.message);
  }

  // ✅ Ambil semua ID Transaksi yang sudah pernah dimasukkan ke sheet "Kirim ke Finance"
  const existingData = targetSheet.getRange(2, 2, targetSheet.getLastRow() - 1, 1).getValues().flat(); // Kolom B
  const existingSet = new Set(existingData.map(String));

  const output = [];
  let no = targetSheet.getLastRow(); // akan lanjut dari row terakhir

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const email = String(row[IDX_EMAIL] || "").trim();
    const resi  = String(row[IDX_FILE_RESI] || "").trim();
    const trx   = String(row[IDX_TRX]).trim();

    if (email.startsWith("Sending completed ✅") && resi === "Ada" && !existingSet.has(trx)) {
      const clean = normalizeWhitespace;
      const nama  = row[IDX_NAMA];
      const tipe  = row[IDX_TIPE];
      const prog  = row[IDX_PROGRAM];
      const stat  = row[IDX_STATUS];

      const fileName = `[Phincon Academy] Bukti Transfer - ${clean(trx)} - ${clean(nama)} - ${clean(tipe)} ${clean(prog)} - ${clean(stat)}`;

      let statusBukti = "Not yet uploaded";
      if (folder.getFilesByName(fileName).hasNext()) {
        statusBukti = "✅ Uploaded";
      }

      output.push([
        no++,
        trx,
        nama,
        tipe,
        prog,
        row[IDX_TANGGAL],
        row[IDX_JUMLAH],
        stat,
        row[IDX_CHANNEL],
        statusBukti,
        ""
      ]);
    }
  }

  // 🧾 Tulis header kalau masih kosong
  if (targetSheet.getLastRow() === 0) {
    targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  }

  // ✅ Tambahkan baris baru tanpa mengganggu existing
  if (output.length > 0) {
    const startRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRow, 1, output.length, outputHeaders.length).setValues(output);
  }

  SpreadsheetApp.getUi().alert("✅ Data berhasil dimuat ke sheet Kirim ke Finance");
}



// // ✅ Fungsi bantu hilangkan spasi berlebih di nama file
// function normalizeWhitespace(str) {
//   return str.toString().replace(/\s+/g, ' ').trim();
// }

// // ✅ Fungsi utama muat ulang data ke sheet "Kirim ke Finance"
// function loadDataKirimKeFinance() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sourceSheet = ss.getSheetByName("Form Responses 1");
//   const targetSheet = ss.getSheetByName("Kirim ke Finance");

//   if (!sourceSheet || !targetSheet) {
//     SpreadsheetApp.getUi().alert("❌ Sheet tidak ditemukan.");
//     return;
//   }

//   const folder = DriveApp.getFolderById(FOLDER_ID_TRANSFER);
//   const sourceData = sourceSheet.getDataRange().getValues();
//   const headerRow = sourceData[0];

//   function findColumnIndex(header, keyword) {
//     return header.findIndex(h => String(h).toLowerCase().includes(keyword.toLowerCase()));
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
//     IDX_TRX, IDX_NAMA, IDX_TIPE, IDX_PROGRAM, IDX_TANGGAL,
//     IDX_JUMLAH, IDX_STATUS, IDX_CHANNEL, IDX_EMAIL, IDX_FILE_RESI
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

//   // 🛡️ Proteksi baris header hanya jika belum ada (Editor tetap bisa jalan)
//   try {
//     const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//     const hasHeader = protections.some(p => p.getRange().getA1Notation() === "A1:K1");
//     if (!hasHeader) {
//       const protect = targetSheet.getRange("A1:K1").protect().setDescription("Header Lock");
//       protect.setWarningOnly(true);
//     }
//   } catch (err) {
//     Logger.log("⚠️ Proteksi header dilewati: " + err.message);
//   }

//   // ⛑ Backup "Status Kirim ke Finance"
//   const existingRows = targetSheet.getLastRow() - 1;
//   const oldStatus = existingRows > 0
//     ? targetSheet.getRange(2, 11, existingRows, 1).getValues()
//     : [];

//   const output = [];
//   let no = existingRows + 1;
//   let index = 0;

//   for (let i = 1; i < sourceData.length; i++) {
//     const row = sourceData[i];
//     const email = String(row[IDX_EMAIL] || "").trim();
//     const resi  = String(row[IDX_FILE_RESI] || "").trim();

//     if (email.startsWith("Sending completed ✅") && resi === "Ada") {
//       const clean = normalizeWhitespace;
//       const trx   = row[IDX_TRX];
//       const nama  = row[IDX_NAMA];
//       const tipe  = row[IDX_TIPE];
//       const prog  = row[IDX_PROGRAM];
//       const stat  = row[IDX_STATUS];

//       const fileName = `[Phincon Academy] Bukti Transfer - ${clean(trx)} - ${clean(nama)} - ${clean(tipe)} ${clean(prog)} - ${clean(stat)}`;
//       let statusBukti = "Not yet uploaded";
//       if (folder.getFilesByName(fileName).hasNext()) {
//         statusBukti = "✅ Uploaded";
//       }

//       const statusKirim = oldStatus[index] ? oldStatus[index][0] : "";
//       index++;

//       output.push([
//         no++,
//         trx,
//         nama,
//         tipe,
//         prog,
//         row[IDX_TANGGAL],
//         row[IDX_JUMLAH],
//         stat,
//         row[IDX_CHANNEL],
//         statusBukti,
//         statusKirim || ""
//       ]);
//     }
//   }

//   // 🧾 Tulis header kalau masih kosong
//   if (targetSheet.getLastRow() === 0) {
//     targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
//   }

//   // ✅ Tambahkan baris baru tanpa mengganggu existing
//   if (output.length > 0) {
//     const startRow = targetSheet.getLastRow() + 1;
//     targetSheet.getRange(startRow, 1, output.length, outputHeaders.length).setValues(output);
//   }

//   SpreadsheetApp.getUi().alert("✅ Data berhasil dimuat ke sheet Kirim ke Finance");
// }