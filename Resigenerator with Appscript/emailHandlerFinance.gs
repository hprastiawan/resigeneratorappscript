// ✅ Fungsi bantu normalisasi spasi
function normalizeWhitespace(str) {
  return str.toString().replace(/\s+/g, ' ').trim();
}

// ✅ Fungsi utama kirim email ke Finance
function sendFinanceEmailFromSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFinance = ss.getSheetByName("Kirim ke Finance");
  const sheetForm = ss.getSheetByName("Form Responses 1");
  const ui = SpreadsheetApp.getUi();

  const folderTransfer = DriveApp.getFolderById(FOLDER_ID_TRANSFER);
  const folderResi = DriveApp.getFolderById(FOLDER_ID_RESI);

  const selection = sheetFinance.getActiveRangeList();
  if (!selection) return;

  const COL_ID_TRX = 2;
  const COL_NAMA = 3;
  const COL_STATUS_BUKTI = 10;

  const formData = sheetForm.getDataRange().getValues();
  const formHeader = formData[0];
  const IDX_ID_FORM = formHeader.findIndex(h => String(h).toLowerCase().includes("id transaksi"));
  const IDX_RESI_STATUS = formHeader.findIndex(h => String(h).toLowerCase().includes("file dalam folder"));

  if (IDX_ID_FORM === -1 || IDX_RESI_STATUS === -1) {
    ui.alert("❌ Kolom ID Transaksi atau File dalam Folder tidak ditemukan di Form Responses 1");
    return;
  }

  const selectedRows = new Set();
  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = start + range.getNumRows() - 1;
    for (let r = start; r <= end; r++) {
      if (r > 1) selectedRows.add(r);
    }
  });

  const rowIndexes = [...selectedRows];
  if (rowIndexes.length < 1) {
    ui.alert("‼️ Pilih minimal 1 baris data terlebih dahulu");
    return;
  }

  // 🔍 Map ID Transaksi => Status Resi dari sheet "Form Responses 1"
  const resiMap = {};
  for (let i = 1; i < formData.length; i++) {
    const id = formData[i][IDX_ID_FORM];
    const status = formData[i][IDX_RESI_STATUS];
    if (id) resiMap[id.toString().trim()] = status;
  }

  const attachments = [];
  const validRows = [];
  const names = [];

  for (const row of rowIndexes) {
    const data = sheetFinance.getRange(row, 1, 1, 11).getValues()[0];

    const idTransaksi = data[1];
    const nama = data[2];
    const tipe = data[3];
    const program = data[4];
    const tanggal = data[5];
    const jumlah = data[6];
    const status = data[7];
    const channel = data[8];
    const statusBukti = data[9];

    const statusResi = resiMap[idTransaksi?.toString().trim()] || "";

    const clean = normalizeWhitespace;
    const buktiName = `[Phincon Academy] Bukti Transfer - ${clean(idTransaksi)} - ${clean(nama)} - ${clean(tipe)} ${clean(program)} - ${clean(status)}`;
    const resiName = `[Receipt Phincon Academy] ${clean(idTransaksi)} - ${clean(nama)} - ${clean(tipe)} ${clean(program)} - ${clean(status)}.pdf`;

    const buktiFiles = folderTransfer.getFilesByName(buktiName);
    const resiFiles = folderResi.getFilesByName(resiName);

    const fileBuktiFound = buktiFiles.hasNext();
    const fileResiFound = resiFiles.hasNext();

    if (
      statusBukti === "✅ Uploaded" &&
      statusResi === "Ada" &&
      fileBuktiFound &&
      fileResiFound
    ) {
      attachments.push(buktiFiles.next().getBlob());
      attachments.push(resiFiles.next().getBlob());

      validRows.push({
        idTransaksi, nama, tipe, program, tanggal, jumlah, status, channel,
        rowIndex: row
      });

      names.push(nama);
    } else {
      sheetFinance.getRange(row, 11).setValue("❌ File tidak lengkap");
      if (!fileBuktiFound) console.warn(`❌ Bukti transfer file NOT FOUND: ${buktiName}`);
      if (!fileResiFound) console.warn(`❌ Resi file NOT FOUND: ${resiName}`);
    }
  }

  if (validRows.length === 0) {
    ui.alert("❌ Tidak ada data valid yang memiliki file resi dan bukti transfer.");
    return;
  }

  const subject = `Dana Masuk Pembayaran Bootcamp - Phincon Academy - ${names.join(" - ")} 🔔`;
  const htmlBody = buildFinanceEmailBody(validRows);

  // ✅ PENGIRIMAN EMAIL KE FINANCE DENGAN TO DAN CC BARU - NEW PROD
  MailApp.sendEmail({
    to: "gloria.daeli@phintraco.com, anisa.sangraningrum@phincon.com, maria.permatasari@phintraco.com",
    subject: subject,
    htmlBody: htmlBody,
    attachments: attachments,
    name: "Phincon Academy System",
    cc: "academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com",
    bcc: "phinconacademy@gmail.com"
  });

  // ✅ Kirim email DEV
  // MailApp.sendEmail({
  //   to: "hendra.prastiawan2@gmail.com", // Versi DEV
  //   subject: subject,
  //   htmlBody: htmlBody,
  //   attachments: attachments,
  //   name: "Phincon Academy System",
  //   cc: "hendra.prastiawan4@gmail.com"
  // });

  validRows.forEach(d => {
    sheetFinance.getRange(d.rowIndex, 11).setValue("✅ Sent to Finance");
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ ${validRows.length} data berhasil dikirim ke Finance`, "📤 Email Finance", 5);
  ui.alert(`✅ Email berhasil dikirim ke Finance untuk ${validRows.length} peserta.`);
}



// ✅ Fungsi bantu membentuk isi email HTML
function buildFinanceEmailBody(dataList) {
  const bannerUrl = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Email%20Header%20Banner%20-%20Phincon%20Academy/Email%20Banner%20Phincon%20Academy%202025.png';
  const logoFooter = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Phincon%20Academy%20-%20Logo%20Footer.png';
  const linkedinIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/linkedin.png';
  const instagramIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/instagram.png';
  const whatsappIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/whatsapp.png';

  const rows = dataList.map(d => `
    <tr>
      <td>${d.idTransaksi}</td>
      <td>${d.nama}</td>
      <td>${d.tipe} ${d.program}</td>
      <td>${d.tanggal}</td>
      <td>${formatRupiah(d.jumlah)}</td>
      <td>${d.channel}</td>
      <td>${d.status}</td>
    </tr>
  `).join('');

  return `
    <div style="font-family: Arial, sans-serif; background-color: #f9f9f9; padding: 30px; max-width: 720px; margin: auto; border-radius: 10px;">
      <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border-radius: 8px; padding: 20px;">
        <tr>
          <td>
            <img src="${bannerUrl}" style="width: 100%; max-width: 680px; border-radius: 8px;" alt="Header Banner Phincon Academy" />
          </td>
        </tr>
        <tr>
          <td style="padding-top: 20px;">
            <h2 style="text-align: center; color: #333;">Notifikasi Dana Masuk</h2>
            <p>Halo Finance Team,</p>
            <p>Berikut adalah data dana masuk untuk pembayaran program bootcamp dari peserta:</p>
            <table width="100%" border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; font-size: 14px;">
              <thead style="background-color: #f3f3f3;">
                <tr>
                  <th>ID Transaksi</th>
                  <th>Nama</th>
                  <th>Program</th>
                  <th>Tanggal</th>
                  <th>Jumlah</th>
                  <th>Channel</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                ${rows}
              </tbody>
            </table>
            <p style="margin-top: 20px;">Terlampir juga bukti transfer dan file resi pembayaran peserta.</p>
            <p style="margin-top: 30px;">Best regards,<br><br><br><strong>Phincon Academy Team</strong></p>
          </td>
        </tr>
        <tr>
          <td style="padding: 30px; background-color: #f3f3f3; text-align: center; border-radius: 8px;">
            <img src="${logoFooter}" alt="Phincon Academy Logo" style="height: 40px; margin-bottom: 10px;" />
            <div style="font-size: 13px; color: #555; line-height: 1.4;">
              Gandaria 8 Office Tower, 8th Floor<br>
              Jl. Arteri Pd. Indah No.10, RT.9/RW.6 Kby. Lama Utara,<br>
              Kec. Kby. Lama, Kota Jakarta Selatan<br>
              Jakarta 12240, Indonesia
            </div>
            <div style="margin-top: 10px;">
              <a href="https://www.linkedin.com/school/phincon-academy/" target="_blank" style="display: inline-block;">
                <img src="${linkedinIcon}" alt="LinkedIn" style="height: 24px; margin: 0 6px;" />
              </a><a href="https://www.instagram.com/phinconacademy" target="_blank" style="display: inline-block;">
                <img src="${instagramIcon}" alt="Instagram" style="height: 24px; margin: 0 6px;" />
              </a><a href="https://api.whatsapp.com/send/?phone=6281119970372" target="_blank" style="display: inline-block;">
                <img src="${whatsappIcon}" alt="WhatsApp" style="height: 24px; margin: 0 6px;" />
              </a>
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;
}

// ✅ Format angka rupiah
function formatRupiah(nominal) {
  if (!nominal || nominal === "-") return "-";
  const number = parseInt(nominal.toString().replace(/[^\d]/g, ''));
  if (isNaN(number)) return "-";
  return `Rp ${number.toLocaleString('id-ID')},-`;
}
