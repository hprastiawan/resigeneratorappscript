function onOpen() {
  const ui = SpreadsheetApp.getUi();
  createMenus(); // Pastikan menu selalu muncul

  // ✅ Jalankan fungsi berat tanpa ganggu menu
  SpreadsheetApp.flush();
  Utilities.sleep(300);

  try { generateStudentId(false); } catch (err) { console.error("generateStudentId error:", err); }
  try { checkDeletedFiles(); } catch (err) { console.error("checkDeletedFiles error:", err); }
  try { applyDropdownNamaProgram(); } catch (err) { console.error("applyDropdownNamaProgram error:", err); }
  // try { applyDropdownLokasi(); } catch (err) { console.error("applyDropdownLokasi error:", err); }
  try { applyDatePickerToTanggal(); } catch (err) { console.error("applyDatePickerToTanggal error:", err); }
  try { updateFormattedTanggal(); } catch (err) { console.error("updateFormattedTanggal error:", err); }
  try { updatePhoneHashing(); } catch (err) { console.error("updatePhoneHashing error:", err); }

  // try { protectFinanceSheetContent(); } catch (err) { console.warn("protectFinanceSheetContent error:", err); }
  try { protectLockedColumns(); } catch (err) { console.warn("protectLockedColumns error:", err); }
  try { validateAndHighlightHeaders(); } catch (err) { console.error("validateAndHighlightHeaders error:", err); }
}

function createMenus() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🏢 Phincon Academy")
    .addItem("🔄 Refresh Halaman", "showRefreshConfirmation")
    //.addItem("🛡 Proteksi Ulang Kolom", "protectLockedColumns")
    .addSeparator()
    .addSubMenu(ui.createMenu("🆔 Generate ID")
      .addItem("🔢 Generate ID Transaksi & Registrasi", "showGenerateIdConfirmation"))
    .addSubMenu(ui.createMenu("🧾 Buat Resi")
      .addItem("📕 Untuk Baris Ini", "showGenerateResiConfirmationforCurrentRow")
      .addItem("📘📗 Untuk Baris Terpilih", "showGenerateResiConfirmationFromSelection")
      .addItem("🚚 Seluruh Data", "showGenerateResiConfirmationForAll"))
    .addSubMenu(ui.createMenu("📧 Kirim Resi ke Email Peserta")
      .addItem("👤 Kirim Resi untuk Baris Ini", "showSendEmailConfirmationForCurrentRow")
      .addItem("👥 Kirim Resi untuk Baris Terpilih", "showSendEmailConfirmationFromSelection")
      .addItem("👥👥 Kirim Resi untuk Seluruh Data", "showSendEmailConfirmationForAll"))
    .addSubMenu(ui.createMenu("📩 Email ke Finance")
      .addItem("🧲 Muat Ulang Data", "showReloadFinanceDataConfirmation")
      .addItem("📤 Upload Bukti Bayar", "openUploadDialogForActiveRow")
      .addItem("🚀 Kirim Email", "showSendFinanceEmailConfirmation"))
    .addToUi();
}

function refreshFileStatus() {
  const ui = SpreadsheetApp.getUi();

  try {
    generateStudentId();
    checkDeletedFiles();
    applyDropdownNamaProgram();
    // applyDropdownLokasi();
    applyDatePickerToTanggal();
    updateFormattedTanggal();
    updatePhoneHashing();

    // try {
    //   protectFinanceSheetContent(); // ✅ Aman, auto-skip jika bukan Owner
    // } catch (err) {
    //   console.warn("⚠️ Gagal proteksi sheet Finance:", err);
    // }

    try {
      protectLockedColumns(); // ✅ Aman, hanya proteksi header baris 1
    } catch (err) {
      console.warn("⚠️ Gagal proteksi kolom:", err);
      ui.showToast("⚠️ Proteksi kolom dilewati");
    }

    createMenus(); // tampilkan ulang menu
    ui.alert("✅ Status file berhasil diperbarui");
  } catch (err) {
    ui.alert("❌ Gagal memperbarui status file");
    console.error("refreshFileStatus error:", err);
  }
}


function protectLockedColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allEditors = ss.getEditors();
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  const ownerEmail = ss.getOwner().getEmail();

  try {
    const headerRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns());

    // Cek apakah sudah ada proteksi di baris header
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .find(p => p.getRange().getA1Notation() === headerRange.getA1Notation());

    if (!existing || currentUserEmail === ownerEmail) {
      let protection = headerRange.protect();
      protection.setDescription("Protect header row");
      protection.removeEditors(protection.getEditors());
      allEditors.forEach(editor => protection.addEditor(editor));
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }
  } catch (err) {
    console.warn(`⚠️ Gagal proteksi baris header: ${err}`);
  }
}



//Hanya owner Spreadsheet yang bisa melakukan editing terhadap kolom-kolom ini
// function protectLockedColumns() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
//   const headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//   const getCol = name => headersRow.indexOf(name);
//   const columnsToProtect = [
//     getCol("ID Transaksi"), getCol("ID Registrasi"), getCol("Otomatis Jangan di edit"),
//     getCol("Nomor Telepon Hashing Otomatis"), getCol("Status Resi PDF"),
//     getCol("File dalam Folder"), getCol("Send Email Status")
//   ].filter(index => index >= 0);

//   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

//   columnsToProtect.forEach(colIndex => {
//     const range = sheet.getRange(2, colIndex + 1, sheet.getMaxRows() - 1);
//     if (!protections.some(p => p.getRange().getA1Notation() === range.getA1Notation())) {
//       const protection = range.protect().setDescription(`Protect column ${colIndex + 1}`);
//       protection.removeEditors(protection.getEditors());
//       if (protection.canDomainEdit()) protection.setDomainEdit(false);
//     }
//   });

//   const headerRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
//   if (!protections.some(p => p.getRange().getA1Notation() === headerRange.getA1Notation())) {
//     const headerProtection = headerRange.protect().setDescription("Protect header row");
//     headerProtection.removeEditors(headerProtection.getEditors());
//     if (headerProtection.canDomainEdit()) headerProtection.setDomainEdit(false);
//   }
// }

function checkDeletedFiles() {
  const { dataRows, headers, folderOutputId, sheet } = getSetup();
  const folder = DriveApp.getFolderById(folderOutputId);

  dataRows.forEach((row, i) => {
    if (!row[headers.trxCol] || !row[headers.regCol]) return;

    const fileName = generateFileName(row, headers);
    const exists = folder.getFilesByName(fileName).hasNext();
    const statusResi = row[headers.statusResiCol];
    const currentStatusFile = row[headers.statusFileCol];

    let statusBaru = "-";
    if (statusResi === "✅ PDF Generated" && exists) statusBaru = "Ada";
    else if (statusResi === "✅ PDF Generated" && !exists) statusBaru = "Pernah dihapus";
    else if (!statusResi || statusResi === "-") statusBaru = "Belum Ada";

    if (currentStatusFile !== statusBaru) {
      const rowInSheet = i + 2;
      sheet.getRange(rowInSheet, headers.statusFileCol + 1).setValue(statusBaru);
    }
  });
}

function validateAndHighlightHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "Form Responses 1") return; // ✅ hanya cek sheet ini saja

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const expectedHeaders = [
    "No", "ID Transaksi", "ID Registrasi", "Email", "Nama Lengkap", "Nomor Telepon",
    "Tipe Program", "Nama Program", "Lokasi", "Status Pembayaran", "Tanggal Transaksi",
    "Otomatis Jangan di edit", "Harga Program", "Down Payment", "Jumlah Pembayaran", "Sisa Pembayaran", "Metode Pembayaran",
    "Cicilan per Bulan (Rp)", "Nomor Telepon Hashing Otomatis", "Channel Pembayaran", "Status Resi PDF", "File dalam Folder", "Send Email Status"
  ];

  for (let i = 0; i < expectedHeaders.length; i++) {
    const actual = (headers[i] || "").toString().trim();
    const expected = expectedHeaders[i];
    const cell = sheet.getRange(1, i + 1);
    cell.setBackground(actual !== expected ? "#f8d7da" : "#d4edda");
  }

  const mismatch = expectedHeaders.some((h, i) => h !== (headers[i] || "").trim());
  if (mismatch) {
    SpreadsheetApp.getUi().alert("⚠️ Urutan header kolom di 'Form Responses 1' tidak sesuai template.\nKolom merah perlu diperiksa");
  }
}


function applyDropdownNamaProgram() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const tipeValues = sheet.getRange(2, 7, sheet.getLastRow() - 1).getValues();
  const programRange = sheet.getRange(2, 8, sheet.getLastRow() - 1);
  const listRTB = getListFromSheet("ListRTB");
  const listPTB = getListFromSheet("ListPTB");
  const combined = [...new Set([...listRTB, ...listPTB])];

  tipeValues.forEach((row, i) => {
    const tipe = (row[0] || "").trim();
    const cell = programRange.getCell(i + 1, 1);
    let rule;

    if (tipe === "RTB") rule = SpreadsheetApp.newDataValidation().requireValueInList(listRTB, true).build();
    else if (tipe === "PTB") rule = SpreadsheetApp.newDataValidation().requireValueInList(listPTB, true).build();
    else rule = SpreadsheetApp.newDataValidation().requireValueInList(combined, true).build();

    cell.setDataValidation(rule);
  });
}

// function applyDropdownLokasi() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
//   const tipeValues = sheet.getRange(2, 7, sheet.getLastRow() - 1).getValues(); // Kolom G (Tipe Program)
//   const lokasiRange = sheet.getRange(2, 9, sheet.getLastRow() - 1); // Kolom I (Lokasi)

//   const listLokasi = getListFromSheet("ListLokasi");
//   const listRTB = listLokasi.filter(l => l.toLowerCase() !== "online");
//   const listPTB = listLokasi.filter(l => l.toLowerCase() === "online");

//   tipeValues.forEach((row, i) => {
//     const tipe = (row[0] || "").trim();
//     const cell = lokasiRange.getCell(i + 1, 1);
//     let rule;

//     if (tipe === "RTB") {
//       rule = SpreadsheetApp.newDataValidation().requireValueInList(listRTB, true).build();
//     } else if (tipe === "PTB") {
//       rule = SpreadsheetApp.newDataValidation().requireValueInList(listPTB, true).build();
//     } else {
//       rule = SpreadsheetApp.newDataValidation().requireValueInList(listLokasi, true).build();
//     }

//     cell.setDataValidation(rule);
//   });
// }

function applyDatePickerToTanggal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const maxRows = sheet.getMaxRows();
  const range = sheet.getRange(2, 11, maxRows - 1); // Kolom K
  const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  range.setDataValidation(rule);
}

function updateFormattedTanggal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const tanggalValues = sheet.getRange(2, 11, lastRow - 1).getValues(); // Kolom K
  const existingFormatted = sheet.getRange(2, 12, lastRow - 1).getValues(); // Kolom L

  const hariIndo = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu"];
  const bulanIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  for (let i = 0; i < tanggalValues.length; i++) {
    const tgl = tanggalValues[i][0];
    const existing = (existingFormatted[i][0] || "").toString().trim();

    const matchJam = existing.match(/(\d{1,2}):(\d{2})$/);
    if (matchJam && matchJam[0] !== "00:00") continue;

    if (!(tgl instanceof Date)) {
      sheet.getRange(i + 2, 12).setValue("");
      continue;
    }

    const hari = hariIndo[tgl.getDay()];
    const tanggal = tgl.getDate();
    const bulan = bulanIndo[tgl.getMonth()];
    const tahun = tgl.getFullYear();
    const formatted = `${hari}, ${tanggal} ${bulan} ${tahun} 00:00`;

    if (formatted !== existing) {
      sheet.getRange(i + 2, 12).setValue(formatted);
    }
  }
}

function updatePhoneHashing() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const lastRow = sheet.getLastRow();
  const phones = sheet.getRange(2, 6, lastRow - 1).getValues(); // Kolom F
  const masked = phones.map(([no]) => {
    const str = no ? no.toString() : "";
    const last3 = str.slice(-3);
    return [str ? "**********" + last3 : ""];
  });
  sheet.getRange(2, 19, masked.length, 1).setValues(masked); // Kolom S
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (sheet.getName() !== "Form Responses 1" || row < 2) return;

  if (col === 7) applyDropdownNamaProgram(); // Kolom G
  if (col === 11) {
    applyDatePickerToTanggal(); // Kolom K
    updateFormattedTanggal();   // Kolom L
  }
  if (col === 6) updatePhoneHashing(); // Kolom F
  if ([7, 8, 9].includes(col)) generateStudentId(); // Kolom G, H, I
}


function getListFromSheet(sheetName) {
  const raw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("A1:A").getValues();
  return raw.flat().filter(v => v && typeof v === "string");
}

// function protectFinanceSheetContent() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Email ke Finance");
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const currentUserEmail = Session.getEffectiveUser().getEmail();
//   const ownerEmail = ss.getOwner().getEmail();

//   // Jalankan hanya kalau yang mengakses adalah Owner
//   if (currentUserEmail !== ownerEmail) {
//     console.warn("🔒 Proteksi sheet Finance dilewati karena kamu bukan owner spreadsheet ini.");
//     return; // Skip untuk Editor
//   }

//   // Hapus semua proteksi yang bisa dihapus
//   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
//   protections.forEach(p => {
//     try {
//       if (p.canEdit()) p.remove();
//     } catch (err) {
//       console.warn(`⚠️ Gagal hapus proteksi lama: ${err}`);
//     }
//   });

//   // Proteksi semua konten (selain header)
//   const lastRow = sheet.getLastRow();
//   const lastCol = sheet.getLastColumn();

//   if (lastRow > 1) {
//     const contentRange = sheet.getRange(2, 1, lastRow - 1, lastCol); // mulai dari baris ke-2
//     const protection = contentRange.protect().setDescription("Protect data isi sheet Finance");
//     protection.removeEditors(protection.getEditors());

//     // Tambahkan semua editor yang diizinkan
//     const editors = ss.getEditors();
//     editors.forEach(editor => protection.addEditor(editor));
//     if (protection.canDomainEdit()) protection.setDomainEdit(false);
//   }
// }


// function protectFinanceSheetContent() {
//   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data Email ke Finance");
//   if (!sheet) return;

//   // Hapus semua proteksi lama
//   const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
//   protections.forEach(p => p.remove());

//   // Proteksi seluruh sheet
//   const protection = sheet.protect().setDescription("Lock sheet except header");
//   protection.setUnprotectedRanges([sheet.getRange("1:1")]); // Hanya header yang tidak dilindungi
//   protection.removeEditors(protection.getEditors());

//   // Jangan biarkan domain edit
//   if (protection.canDomainEdit()) protection.setDomainEdit(false);
// }
