// ✅ Kirim email hanya untuk baris aktif (1 row)
function sendEmailForCurrentRow() {
  const { sheet, headers, folderOutputId } = getSetup();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();

  if (row === 1) {
    ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const email = rowData[headers.emailCol];
  const trxId = rowData[headers.trxCol];
  const regId = rowData[headers.regCol];
  const name = rowData[headers.nameCol];

  if (!email || !trxId || !regId) {
    ui.alert("❌ Data belum lengkap. Pastikan Email, ID Transaksi, dan ID Registrasi sudah terisi");
    return;
  }

  const folder = DriveApp.getFolderById(folderOutputId);
  const fileName = generateFileName(rowData, headers);
  const files = folder.getFilesByName(fileName);

  if (!files.hasNext()) {
    sheet.getRange(row, headers.sendEmailStatusCol + 1).setValue("Sending failed ❌: File Resi tidak ada di GDrive");
    ui.alert(`🚫 File resi "${fileName}" tidak ditemukan di folder`);
    return;
  }

  try {
    const pdf = files.next().getAs(MimeType.PDF);
    sendConfirmationEmail(rowData, headers, pdf);
    sheet.getRange(row, headers.sendEmailStatusCol + 1).setValue("Sending completed ✅");
    ui.alert(`✅ Email berhasil dikirim ke ${name}`);
  } catch (err) {
    const msg = err.message || "Unknown error";
    sheet.getRange(row, headers.sendEmailStatusCol + 1).setValue("Sending failed ❌: " + msg);
    ui.alert(`❌ Gagal mengirim email: ${msg}`);
  }
}


// ✅ Handler: Kirim Email → Untuk Baris Terpilih
function sendEmailForSelection() {
  const { sheet, headers, folderOutputId } = getSetup();
  const selection = sheet.getActiveRangeList();

  if (!selection) {
    SpreadsheetApp.getUi().alert("⛔️ Tidak ada baris yang dipilih");
    return;
  }

  const folder = DriveApp.getFolderById(folderOutputId);
  const ranges = selection.getRanges();
  let processed = 0;
  let success = 0;
  let failed = 0;

  for (const range of ranges) {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const values = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

    for (let i = 0; i < values.length; i++) {
      const rowIndex = startRow + i;
      const row = values[i];
      const name = row[headers.nameCol];
      const email = row[headers.emailCol];
      const trxId = row[headers.trxCol];
      const regId = row[headers.regCol];

      SpreadsheetApp.getActive().toast(`📨 Mengirim ke ${name} (${processed + 1})...`);
      Utilities.sleep(300); // jeda antar kiriman

      if (!email || !trxId || !regId) {
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: Data tidak lengkap");
        failed++;
        processed++;
        continue;
      }

      const fileName = generateFileName(row, headers);
      const files = folder.getFilesByName(fileName);

      if (!files.hasNext()) {
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: File tidak ditemukan");
        failed++;
        processed++;
        continue;
      }

      try {
        const pdf = files.next().getAs(MimeType.PDF);
        sendConfirmationEmail(row, headers, pdf);
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending completed ✅");
        success++;
      } catch (err) {
        const msg = err.message || "Unknown error";
        sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
          .setValue("Sending failed ❌: " + msg);
        failed++;
      }

      processed++;
    }
  }

  // ✅ Dialog summary dengan checkbox icon dan tombol OK
  SpreadsheetApp.getUi().alert(
    `✅ ${success} email berhasil dikirim\n❌ ${failed} gagal dikirim\n📦 Total diproses: ${processed} baris`
  );
}

// ✅ Untuk seluruh data — tampilkan summary dalam modal alert
function sendEmailToAllRows() {
  const { sheet, headers, folderOutputId } = getSetup();
  const folder = DriveApp.getFolderById(folderOutputId);
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  let processed = 0;
  let success = 0;
  let failed = 0;

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    const row = values[i];
    const name = row[headers.nameCol];
    const email = row[headers.emailCol];
    const trxId = row[headers.trxCol];
    const regId = row[headers.regCol];

    SpreadsheetApp.getActive().toast(`📨 Mengirim ke ${name} (${processed + 1})...`);
    Utilities.sleep(300);

    if (!email || !trxId || !regId) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: Data tidak lengkap");
      failed++;
      processed++;
      continue;
    }

    const fileName = generateFileName(row, headers);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: File tidak ditemukan");
      failed++;
      processed++;
      continue;
    }

    try {
      const pdf = files.next().getAs(MimeType.PDF);
      sendConfirmationEmail(row, headers, pdf);
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending completed ✅");
      success++;
    } catch (err) {
      const msg = err.message || "Unknown error";
      sheet.getRange(rowIndex, headers.sendEmailStatusCol + 1)
        .setValue("Sending failed ❌: " + msg);
      failed++;
    }

    processed++;
  }

  // ✅ Ringkasan akhir dengan ikon dan tombol OK
  SpreadsheetApp.getUi().alert(
    `✅ ${success} email berhasil dikirim\n❌ ${failed} gagal dikirim\n📦 Total diproses: ${processed} baris`
  );
}
