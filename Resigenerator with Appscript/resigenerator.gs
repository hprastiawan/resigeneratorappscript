function generateStudentId(showAlert = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  const values = sheet.getDataRange().getValues();

  const rtbProgramCodes = getProgramCodesFromSheet("ListRTB");
  const ptbProgramCodes = getProgramCodesFromSheet("ListPTB");

  const locationMapAlpha = { "Jakarta": "J", "Yogya": "Y", "Online": "O" };
  const locationMapNumeric = { "Jakarta": "1", "Yogya": "2", "Online": "3" };

  let regChanges = [];
  let txChanges = [];
  let newlyGenerated = 0;

  for (let i = 1; i < values.length; i++) {
    const row = i + 1;
    const rowData = values[i];

    const fullName = rowData[4];
    const phoneNumber = rowData[5];
    const typeProgram = rowData[6];
    const programName = rowData[7];
    const location = rowData[8];
    const rawDate = rowData[10];

    const existingTxId = rowData[1] ? rowData[1].toString().trim() : "";
    const existingRegId = rowData[2] ? rowData[2].toString().trim() : "";

    const idTransaksiCell = sheet.getRange(row, 2);
    const idRegistrasiCell = sheet.getRange(row, 3);
    const dateCell = sheet.getRange(row, 11); // K
    const dateTextCell = sheet.getRange(row, 12); // L
    const phoneHashCell = sheet.getRange(row, 19); // S

    // 🔐 Hash nomor telepon
    phoneHashCell.setValue(hashPhoneNumber(phoneNumber));

    // 📆 Format tanggal
    if (rawDate instanceof Date) {
      const copiedDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
      dateCell.setValue(copiedDate);

      const currentText = rowData[11] ? rowData[11].toString().trim() : "";
      const matchJam = currentText.match(/(\d{1,2}):(\d{2})$/);
      const hasManualTime = matchJam && matchJam[0] !== "00:00";

      if (!hasManualTime) {
        dateTextCell.setValue(formatIndonesianDate(copiedDate));
      }
    }

    // 📋 Validasi dropdown Nama Program (H)
    const rangeH = sheet.getRange(row, 8);
    if (typeProgram === 'RTB') {
      const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListRTB').getRange('A1:A').getValues().filter(String);
      rangeH.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list.flat(), true).build());
    } else if (typeProgram === 'PTB') {
      const list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListPTB').getRange('A1:A').getValues().filter(String);
      rangeH.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list.flat(), true).build());
    } else {
      rangeH.clearDataValidations();
    }

    // 🧠 Generate prefix
    const regPrefix = `${locationMapAlpha[location?.trim()] || "X"}${typeProgram === 'RTB' ? "1" : "2"}`;
    const txPrefix = `${locationMapNumeric[location?.trim()] || "X"}${typeProgram === 'RTB' ? "1" : "2"}`;
    const programCode = typeProgram === 'RTB' ? rtbProgramCodes[programName?.trim()] ?? "000" : ptbProgramCodes[programName?.trim()] ?? "000";

    // 🆔 ID REGISTRASI
    const expectedRegPrefix = `${regPrefix}${programCode}`;
    const shouldRegenerateReg = !existingRegId || !existingRegId.startsWith(expectedRegPrefix);

    if (fullName && typeProgram && programName && location && shouldRegenerateReg) {
      let finalRegId;
      if (existingRegId && existingRegId.length >= 4) {
        const suffix = existingRegId.slice(-4);
        finalRegId = `${expectedRegPrefix}${suffix}`;
      } else {
        finalRegId = `${expectedRegPrefix}${generateRandomCode(4)}`;
      }

      if (!existingRegId) {
        idRegistrasiCell.setValue(finalRegId);
        newlyGenerated++;
      } else if (existingRegId !== finalRegId) {
        idRegistrasiCell.setValue(finalRegId);
        regChanges.push(fullName);
      }
    }

    // 🧾 ID TRANSAKSI
    if (rawDate instanceof Date && typeProgram && programName && location) {
      const yy = String(rawDate.getFullYear()).slice(-2);
      const mm = String(rawDate.getMonth() + 1).padStart(2, '0');
      const dd = String(rawDate.getDate()).padStart(2, '0');
      const dateKey = `${yy}${mm}${dd}`;
      const txPrefixFull = `${txPrefix}${programCode}${dateKey}`;

      // Hitung jumlah baris sebelumnya dengan kombinasi yang sama
      let countSameBefore = 0;
      for (let j = 1; j < i; j++) {
        const prevRow = values[j];
        const [prevType, prevProg, prevLoc, prevDate] = [prevRow[6], prevRow[7], prevRow[8], prevRow[10]];

        if (
          prevType === typeProgram &&
          prevProg === programName &&
          prevLoc === location &&
          prevDate instanceof Date &&
          prevDate.getFullYear() === rawDate.getFullYear() &&
          prevDate.getMonth() === rawDate.getMonth() &&
          prevDate.getDate() === rawDate.getDate()
        ) {
          countSameBefore++;
        }
      }

      const order = String(countSameBefore + 1).padStart(3, '0');
      const finalTxId = `${txPrefixFull}${order}`;

      if (!existingTxId) {
        idTransaksiCell.setValue(finalTxId);
        newlyGenerated++;
      } else if (existingTxId !== finalTxId) {
        idTransaksiCell.setValue(finalTxId);
        txChanges.push(fullName);
      }
    }
  }

  // 📅 Re-apply date picker
  const lastRow = sheet.getLastRow();
  const dateRange = sheet.getRange(2, 11, lastRow - 1);
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  dateRange.setDataValidation(dateRule);

  // 🔔 ALERT SESUAI EKSPEKTASI
  if (showAlert) {
    let alertMsg = "";

    if (newlyGenerated > 0 && regChanges.length === 0 && txChanges.length === 0) {
      alertMsg = "✅ ID Transaksi & Registrasi berhasil dibuat dan diupdate";
    } else if (regChanges.length > 0 || txChanges.length > 0) {
      const list = [];
      regChanges.forEach(n => list.push(`✅ ID Registrasi atas nama ${n} mengalami perubahan`));
      txChanges.forEach(n => list.push(`✅ ID Transaksi atas nama ${n} mengalami perubahan`));
      alertMsg = list.join("\n");
    } else {
      alertMsg = "ℹ️ ID Transaksi / ID Registrasi tidak ada perubahan dari data yang ada saat ini";
    }

    SpreadsheetApp.getUi().alert(alertMsg);
  }
}

function generateRandomCode(length = 4) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  return Array.from({ length }, () => chars[Math.floor(Math.random() * chars.length)]).join("");
}

function formatIndonesianDate(date) {
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
    'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return `${days[date.getDay()]}, ${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()} 00:00`;
}

function hashPhoneNumber(phone) {
  if (!phone) return '';
  const digits = phone.toString().replace(/\D/g, '');
  if (digits.length <= 3) return 'x'.repeat(digits.length);
  return 'x'.repeat(digits.length - 3) + digits.slice(-3);
}

function getProgramCodesFromSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {};

  const data = sheet.getRange("A1:A").getValues().flat().filter(Boolean);
  const total = data.length;
  const codeLength = total >= 100 ? 3 : (total >= 10 ? 2 : 1);

  const programCodes = {};
  data.forEach((name, i) => {
    const code = String(i + 1).padStart(codeLength, "0");
    programCodes[name.trim()] = code;
  });

  return programCodes;
}
