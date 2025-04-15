// ✅ Fungsi utama: generate semua data
function generateResiPDF() {
  const { dataRows, headers, folderOutputId, slideTemplateId, sheet } = getSetup();
  const ui = SpreadsheetApp.getUi();
  const outputFolder = DriveApp.getFolderById(folderOutputId);

  const total = dataRows.filter(row => row[headers.trxCol] && row[headers.regCol]).length;
  let count = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (!row[headers.trxCol] || !row[headers.regCol]) continue;

    const fileName = generateFileName(row, headers);
    const existing = outputFolder.getFilesByName(fileName);

    const statusResi = row[headers.statusResiCol];
    const fileStatusCol = headers.statusFileCol + 1;

    // ✅ Tentukan status file
    let fileStatus = "-";
    if (statusResi === "✅ PDF Generated" && existing.hasNext()) {
      fileStatus = "Ada";
      sheet.getRange(i + 2, fileStatusCol).setValue(fileStatus);
      continue;
    } else if (statusResi === "✅ PDF Generated" && !existing.hasNext()) {
      fileStatus = "Pernah dihapus";
    } else {
      fileStatus = "Belum Ada";
    }
    sheet.getRange(i + 2, fileStatusCol).setValue(fileStatus);

    if (fileStatus !== "Pernah dihapus" && fileStatus !== "Belum Ada") continue;

    headers.rowIndex = i;
    createResiPDF(row, headers, slideTemplateId, outputFolder, fileName, sheet);
    count++;
    SpreadsheetApp.getActiveSpreadsheet().toast(`${count} of ${total} files generated...`, "Progress", 3);
  }

  ui.alert(`✅ Semua resi berhasil digenerate (${count} file baru).`);
}


// ✅ Fungsi: generate hanya baris aktif
function generateResiPDFforCurrentRow() {
  const { sheet, headers, slideTemplateId, folderOutputId } = getSetup();
  const ui = SpreadsheetApp.getUi();
  const row = sheet.getActiveRange().getRow();
  if (row === 1) return ui.alert("⛔️ Baris judul (header) tidak boleh dipilih");

  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!data[headers.trxCol] || !data[headers.regCol]) {
    return ui.alert("🚫 Data belum lengkap untuk baris ini");
  }

  headers.rowIndex = row - 2;
  const fileName = generateFileName(data, headers);
  const outputFolder = DriveApp.getFolderById(folderOutputId);
  const existing = outputFolder.getFilesByName(fileName);

  if (existing.hasNext()) {
    const response = ui.alert(`❗File \"${fileName}\" sudah ada. Mau ganti?`, ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) return;
    existing.next().setTrashed(true);
  }

  createResiPDF(data, headers, slideTemplateId, outputFolder, fileName, sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(`1 file berhasil digenerate`, "Progress", 3);
  ui.alert(`✅ File berhasil digenerate untuk: ${data[headers.nameCol]}`);
}


// ✅ Fungsi: generate berdasarkan baris yang di-block
function generateResiPDFFromSelection() {
  const { sheet, headers, slideTemplateId, folderOutputId } = getSetup();
  const ui = SpreadsheetApp.getUi();
  const selection = sheet.getActiveRange();
  const outputFolder = DriveApp.getFolderById(folderOutputId);

  const dataRows = selection.getValues();
  const total = dataRows.filter(r => r[headers.trxCol] && r[headers.regCol]).length;
  let count = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    if (!row[headers.trxCol] || !row[headers.regCol]) continue;

    headers.rowIndex = selection.getRow() - 2 + i;
    const fileName = generateFileName(row, headers);
    const existing = outputFolder.getFilesByName(fileName);

    // ✅ Update kolom "File dalam Folder"
    let fileStatus = "-";
    if (existing.hasNext()) {
      fileStatus = "Ada";
    } else {
      fileStatus = row[headers.statusResiCol] === "✅ PDF Generated" ? "Pernah dihapus" : "Belum Ada";
    }
    sheet.getRange(headers.rowIndex + 2, headers.statusFileCol + 1).setValue(fileStatus);

    // ✅ Hanya generate ulang jika "Pernah dihapus" atau "Belum Ada"
    if (fileStatus === "Pernah dihapus" || fileStatus === "Belum Ada") {
      createResiPDF(row, headers, slideTemplateId, outputFolder, fileName, sheet);
      count++;
      SpreadsheetApp.getActiveSpreadsheet().toast(`${count} of ${total} files generated...`, "Progress", 3);
    }
  }

  ui.alert(`✅ ${count} file resi berhasil digenerate dari baris terpilih`);
}


// ✅ Fungsi bantu: menyimpan setup awal
function getSetup() {
  const slideTemplateId = '1QGHf2v9-9xyYFyHKYhHdbs6gmOC-7JiwELePAVIMS_U';
  const folderOutputId = '1XnMUkIPdzJpDZGL_F3wOw2Tu7JaJUuj-';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const data = sheet.getDataRange().getDisplayValues(); // ✅ AMBIL DISPLAY

  const headersRow = data[0];
  const getCol = name => headersRow.indexOf(name);

  const headers = {
    trxCol: getCol("ID Transaksi"),
    regCol: getCol("ID Registrasi"),
    emailCol: getCol("Email"),
    nameCol: getCol("Nama Lengkap"),
    phoneCol: getCol("Nomor Telepon Hashing Otomatis"),
    tglTextCol: getCol("Otomatis Jangan di edit"),
    channelCol: getCol("Channel Pembayaran"),
    tipeCol: getCol("Tipe Program"),
    programCol: getCol("Nama Program"),
    statusCol: getCol("Status Pembayaran"),
    hargaCol: getCol("Harga Program"),
    jmlBayarCol: getCol("Jumlah Pembayaran"),
    diskonCol: getCol("Diskon Program"),
    dpCol: getCol("Down Payment"),
    sisaCol: getCol("Sisa Pembayaran"),
    metodeCol: getCol("Metode Pembayaran"),
    cicilanCol: getCol("Cicilan per Bulan (Rp)"),
    statusResiCol: getCol("Status Resi PDF"),
    statusFileCol: getCol("File dalam Folder"),
    sendEmailStatusCol: getCol("Send Email Status")
  };

  return { dataRows: data.slice(1), headers, folderOutputId, slideTemplateId, sheet };
}



// ✅ Fungsi bantu: generate nama file PDF
function generateFileName(row, h) {
  return `[Receipt Phincon Academy] ${row[h.trxCol]} - ${row[h.nameCol]} - ${row[h.tipeCol]} ${row[h.programCol]} - ${row[h.statusCol]}.pdf`.trim();
}


// ✅ Fungsi bantu: generate isi file resi
function createResiPDF(row, h, slideTemplateId, outputFolder, fileName, sheet) {
  const slideCopy = DriveApp.getFileById(slideTemplateId).makeCopy(`Resi - ${row[h.nameCol]}`);
  const presentation = SlidesApp.openById(slideCopy.getId());
  const slide = presentation.getSlides()[0];

  // Fungsi internal untuk format angka jadi 9.500.000,-
  const formatManual = val => {
    if (!val || val === "-" || val === "0" || val === "Rp 0") return "-";
    const cleaned = val.toString().replace(/[^0-9]/g, '');
    const formatted = cleaned.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return `${formatted},-`;
  };

  const replacements = {
    '<<tanggaltrx>>': row[h.tglTextCol],
    '<<idtrx>>': row[h.trxCol],
    '<<channel>>': row[h.channelCol],
    '<<idreg>>': row[h.regCol],
    '<<namapeserta>>': row[h.nameCol],
    '<<email>>': row[h.emailCol],
    '<<notlp>>': row[h.phoneCol],
    '<<tipeprog>>': row[h.tipeCol],
    '<<namaprog>>': row[h.programCol],
    '<<statusbayar>>': row[h.statusCol],

    '<<jmlbayar>>': formatManual(row[h.jmlBayarCol]),
    '<<hargaprog>>': formatManual(row[h.hargaCol]),
    '<<diskonprog>>': formatManual(row[h.diskonCol]),
    '<<dpprog>>': formatManual(row[h.dpCol]),
    '<<sisabayar>>': formatManual(row[h.sisaCol]),
    '<<cicilanperbulanrp>>': formatManual(row[h.cicilanCol]),

    '<<metodebayar>>': row[h.metodeCol],
    '<<rp>>': row[h.cicilanCol] && row[h.cicilanCol] !== "-" && row[h.cicilanCol] !== "0" ? "Rp." : ""
  };

  for (const [tag, value] of Object.entries(replacements)) {
    slide.replaceAllText(tag, value);
  }

  presentation.saveAndClose();
  const blob = DriveApp.getFileById(slideCopy.getId()).getAs(MimeType.PDF);
  blob.setName(fileName);
  outputFolder.createFile(blob);
  DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

  if (typeof h.rowIndex !== 'undefined') {
    const rowInSheet = h.rowIndex + 2;
    sheet.getRange(rowInSheet, h.statusResiCol + 1).setValue("✅ PDF Generated");
    sheet.getRange(rowInSheet, h.statusFileCol + 1).setValue("Ada");
  }
}

function formatCicilan(nominal) {
  if (!nominal || nominal === "-") return "-";
  const formatted = formatRupiah(nominal);
  return `Rp.\t${formatted}`;
}
