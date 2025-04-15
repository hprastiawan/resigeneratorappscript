// ✅ Konfirmasi Menu untuk Refresh Halaman
function showRefreshConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", "Apakah Kamu yakin ingin me-refresh halaman ini?", ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    refreshFileStatus();
  }
}

// ✅ Konfirmasi Menu untuk Generate ID
function showGenerateIdConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", "Apakah Kamu yakin ingin generate ID Transaksi & Registrasi?", ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateStudentId(true);
  }
}

// ✅ Konfirmasi Menu Cetak Resi → Untuk Baris Ini
function showGenerateResiConfirmationforCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRow = sheet.getActiveCell().getRow();

  // ✅ Cek apakah baris header (baris 1) dipilih
  if (activeRow === 1) {
    SpreadsheetApp.getUi().alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ✅ Cek jika belum memilih baris valid
  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert("‼️ Pilih salah satu baris data terlebih dahulu");
    return;
  }

  const nama = sheet.getRange(activeRow, 5).getValue(); // Kolom E = Nama Lengkap
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", `Apakah Kamu yakin ingin membuat resi untuk ${nama}?`, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateResiPDFforCurrentRow();
  }
}


// ✅ Konfirmasi Menu Cetak Resi → Untuk Baris Terpilih
function showGenerateResiConfirmationFromSelection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();
  if (!selection) return;

  const selectedRows = new Set();
  let headerSelected = false;

  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i === 1) headerSelected = true;
      if (i >= 2) selectedRows.add(i);
    }
  });

  if (headerSelected) {
    SpreadsheetApp.getUi().alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 2) {
    SpreadsheetApp.getUi().alert("‼️ Pilih minimal 2 baris data terlebih dahulu");
    return;
  }

  const namaList = rowIndexes.map(row => sheet.getRange(row, 5).getValue()).filter(n => n).join(", ");
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", `Apakah Kamu yakin ingin membuat resi untuk ${rowIndexes.length} peserta berikut?\n\n${namaList}`, ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateResiPDFFromSelection();
  }
}

// ✅ Konfirmasi Menu Cetak Resi → Seluruh Data
function showGenerateResiConfirmationForAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("⚠️ Konfirmasi", "🚨 Apakah Kamu yakin ingin membuat Resi untuk seluruh data yang ada?", ui.ButtonSet.OK_CANCEL);

  if (response === ui.Button.OK) {
    generateResiPDF();
  }
}

// ✅ Konfirmasi Menu Kirim Email → Untuk Baris Ini
function showSendEmailConfirmationForCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRow = sheet.getActiveCell().getRow();
  const ui = SpreadsheetApp.getUi();

  // ⛔️ Cek baris judul
  if (activeRow === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  // ⛔️ Cek jika belum pilih baris valid
  if (activeRow < 2) {
    ui.alert("⛔️ Pilih salah satu baris data terlebih dahulu");
    return;
  }

  const email = sheet.getRange(activeRow, 4).getValue(); // Kolom D = Email
  const nama = sheet.getRange(activeRow, 5).getValue();  // Kolom E = Nama Lengkap
  const trxId = sheet.getRange(activeRow, 2).getValue(); // Kolom B = ID Transaksi
  const regId = sheet.getRange(activeRow, 3).getValue(); // Kolom C = ID Registrasi

  // ⛔️ Validasi data wajib
  if (!email || !trxId || !regId) {
    ui.alert("❌ Data belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi");
    return;
  }

  const response = ui.alert(
    "⚠️ Konfirmasi",
    `🚨 Apakah Kamu yakin ingin mengirim resi ke email peserta berikut?\n\n👤 Nama: ${nama}\n📧 Email: ${email}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    sendEmailForCurrentRow();
  }
}

// ✅ Konfirmasi Menu Kirim Email → Untuk Baris Terpilih
function showSendEmailConfirmationFromSelection() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRangeList();
  const ui = SpreadsheetApp.getUi();

  if (!selection) return;

  const selectedRows = new Set();
  let headerSelected = false;

  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i === 1) headerSelected = true;
      if (i >= 2) selectedRows.add(i);
    }
  });

  if (headerSelected) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 2) {
    ui.alert("‼️ Pilih minimal 2 baris data terlebih dahulu");
    return;
  }

  // Validasi data wajib (Email, ID Transaksi, ID Registrasi)
  const incomplete = rowIndexes.filter(row => {
    const email = sheet.getRange(row, 4).getValue(); // Kolom D
    const trxId = sheet.getRange(row, 2).getValue(); // Kolom B
    const regId = sheet.getRange(row, 3).getValue(); // Kolom C
    return !(email && trxId && regId);
  });

  if (incomplete.length > 0) {
    ui.alert("❌ Beberapa baris belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi di semua baris yang dipilih");
    return;
  }

  const namaList = rowIndexes.map(row => sheet.getRange(row, 5).getValue()).filter(n => n).join(", ");
  const response = ui.alert(
    "⚠️ Konfirmasi",
    `🚨 Apakah Kamu yakin ingin mengirim resi ke email untuk ${rowIndexes.length} peserta berikut?\n\n${namaList}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    sendEmailForSelection();
  }
}

// ✅ Konfirmasi Menu Kirim Resi → Seluruh Data
function showSendEmailConfirmationForAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ Konfirmasi",
    "🚨 Apakah Kamu yakin ingin mengirim resi ke seluruh peserta dalam data?",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    sendEmailToAllRows();
  }
}

// ✅ Konfirmasi Menu Email ke Finance → Muat Ulang Data
function showReloadFinanceDataConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ Konfirmasi",
    "Apakah Kamu yakin ingin memuat ulang data Finance?\n\nLangkah ini akan memperbarui data yang ada saat ini",
    ui.ButtonSet.OK_CANCEL
  );

  if (response === ui.Button.OK) {
    //loadFinanceData();
    loadDataKirimKeFinance();
  }
}

function showSendFinanceEmailConfirmation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFinance = ss.getSheetByName("Kirim ke Finance");
  const sheetForm = ss.getSheetByName("Form Responses 1");
  const ui = SpreadsheetApp.getUi();

  if (!sheetFinance || !sheetForm) {
    ui.alert("❌ Sheet tidak ditemukan.");
    return;
  }

  // Kolom untuk sheet Data Email ke Finance
  const COL_NAMA = 3;
  const COL_ID_TRX = 2;
  const COL_STATUS_BUKTI = 10;

  // Kolom untuk Form Responses 1 (khusus validasi resi)
  const dataForm = sheetForm.getDataRange().getValues();
  const headerForm = dataForm[0];
  const IDX_ID_TRX_FORM = headerForm.findIndex(h => String(h).toLowerCase().includes("id transaksi"));
  const IDX_RESI_STATUS = headerForm.findIndex(h => String(h).toLowerCase().includes("file dalam folder"));

  if (IDX_ID_TRX_FORM === -1 || IDX_RESI_STATUS === -1) {
    ui.alert("❌ Kolom 'ID Transaksi' atau 'File dalam Folder' tidak ditemukan di Form Responses.");
    return;
  }

  // Ambil selection
  const selection = sheetFinance.getActiveRangeList();
  if (!selection) return;

  const selectedRows = new Set();
  let headerSelected = false;

  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let i = start; i <= end; i++) {
      if (i === 1) headerSelected = true;
      if (i >= 2) selectedRows.add(i);
    }
  });

  if (headerSelected) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 1) {
    ui.alert("‼️ Pilih minimal 1 baris data terlebih dahulu");
    return;
  }

  // 🔍 Buat mapping ID Transaksi → Status Resi dari Form Responses 1
  const resiMap = {};
  for (let i = 1; i < dataForm.length; i++) {
    const row = dataForm[i];
    const id = row[IDX_ID_TRX_FORM];
    const statusResi = row[IDX_RESI_STATUS];
    if (id) resiMap[id.toString().trim()] = statusResi;
  }

  // Validasi kelengkapan
  const missingFiles = rowIndexes.filter(row => {
    const idTransaksi = sheetFinance.getRange(row, COL_ID_TRX).getValue().toString().trim();
    const statusResi = resiMap[idTransaksi] || "";
    const statusBukti = sheetFinance.getRange(row, COL_STATUS_BUKTI).getValue();
    return statusResi !== "Ada" || statusBukti !== "✅ Uploaded";
  });

  if (missingFiles.length > 0) {
    ui.alert("❌ Beberapa data tidak lengkap. Pastikan file resi dan bukti transfer tersedia untuk semua baris yang dipilih.");
    return;
  }

  if (rowIndexes.length > 1) {
    ui.alert("✅ Semua data sudah lengkap. Akan dilanjutkan ke konfirmasi pengiriman.");
  }

  try {
    const namaList = rowIndexes
      .map(row => sheetFinance.getRange(row, COL_NAMA).getValue())
      .filter(n => n)
      .join(", ");

    const response = ui.alert(
      "⚠️ Konfirmasi",
      `🚨 Apakah Kamu yakin ingin mengirim email ke Finance untuk ${rowIndexes.length} peserta berikut?\n\n${namaList}`,
      ui.ButtonSet.OK_CANCEL
    );

    if (response === ui.Button.OK) {
      sendFinanceEmailFromSelection();
    }
  } catch (e) {
    ui.alert("❌ Terjadi error saat membaca data. Silakan coba ulangi.");
  }
}