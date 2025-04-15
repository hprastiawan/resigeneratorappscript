// ✅ Template email dan attachment
function sendConfirmationEmail(row, h, blob) {
  const email = row[h.emailCol];
  const name = row[h.nameCol];
  const trxId = row[h.trxCol];
  const tgl = row[h.tglTextCol];
  const channel = row[h.channelCol];
  const jumlah = row[h.jmlBayarCol];
  const tipe = row[h.tipeCol];
  const program = row[h.programCol];
  const status = row[h.statusCol];

  const bannerUrl = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Email%20Header%20Banner%20-%20Phincon%20Academy/Email%20Banner%20Phincon%20Academy%202025.png';
  const logoFooter = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Phincon%20Academy%20-%20Logo%20Footer.png';
  const linkedinIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/linkedin.png';
  const instagramIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/instagram.png';
  const whatsappIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/whatsapp.png';

  const rawSubject = `[Phincon Academy] ${trxId} - ${name} - ${tipe} ${program} - ${status} - Berhasil diterima 🎉`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; background-color: #f9f9f9 !important; padding: 30px; max-width: 640px; margin: auto; border-radius: 10px;">
      <table width="100%" cellspacing="0" cellpadding="0" bgcolor="#ffffff" style="background-color: #ffffff !important; padding: 20px; border-radius: 8px; color: #333 !important;">
        <tr>
          <td>
            <img src="${bannerUrl}" alt="Phincon Academy Banner" style="width: 100%; max-width: 600px; height: auto; border-radius: 8px;" />
          </td>
        </tr>
        <tr>
          <td style="padding-top: 20px;">
            <h2 style="color: #333; text-align: center;">Pembayaran Kamu Telah Berhasil</h2>
            <p>Hi <strong>${name}</strong>,</p>
            <p>Selamat, pembayaran kamu sudah berhasil kami terima.</p>
            <p style="color: #21a366; font-weight: bold;">Detail Pembayaran</p>
            <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse: collapse; font-size: 14px;">
              <tr style="background-color: #f3fdf6;">
                <td><strong>ID Transaksi</strong></td>
                <td><strong>${trxId}</strong></td>
              </tr>
              <tr>
                <td>Nama Program</td>
                <td>${tipe} ${program}</td>
              </tr>
              <tr>
                <td>Channel Pembayaran</td>
                <td>${channel}</td>
              </tr>
              <tr>
                <td>Tanggal Transaksi</td>
                <td>${tgl}</td>
              </tr>
              <tr>
                <td>Jumlah</td>
                <td><strong>${formatRupiah(jumlah)}</strong></td>
              </tr>
              <tr>
                <td>Status Pembayaran</td>
                <td><strong>${status}</strong></td>
              </tr>
            </table>
            <p style="margin-top: 16px;">Silakan temukan bukti pembayaran Kamu pada lampiran email ini.</p>
            <p>Terima kasih atas kepercayaan Kamu kepada Phincon Academy.</p>
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
                <img src="${linkedinIcon}" style="height: 24px; margin: 0 6px;" />
              </a><a href="https://www.instagram.com/phinconacademy?igsh=MzRlODBiNWFlZA%3D%3D" target="_blank" style="display: inline-block;">
                <img src="${instagramIcon}" style="height: 24px; margin: 0 6px;" />
              </a><a href="https://api.whatsapp.com/send/?phone=6281119970372&text&type=phone_number&app_absent=0" target="_blank" style="display: inline-block;">
                <img src="${whatsappIcon}" style="height: 24px; margin: 0 6px;" />
              </a>
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;

  //PROD
  MailApp.sendEmail({
  to: email,
  subject: rawSubject,
  htmlBody: htmlBody,
  attachments: [blob],
  name: 'Phincon Academy',
  //bcc: 'hendra.prastiawan4@gmail.com'
  bcc: 'academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com, phinconacademy@gmail.com'
});

//   //DEV NEW
//   MailApp.sendEmail({
//   to: email,
//   subject: rawSubject,
//   htmlBody: htmlBody,
//   attachments: [blob],
//   name: 'Phincon Academy',
//   bcc: 'hendra.prastiawan4@gmail.com, phinconacademy@gmail.com'
//   //bcc: 'academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com'
// });

  // GmailApp.sendEmail(email, encodeSubjectUTF8(rawSubject), '', {
  //   htmlBody: htmlBody,
  //   attachments: [blob],
  //   name: 'Phincon Academy',
  //   bcc: 'hendra.prastiawan4@gmail.com'
  //   //bcc: 'academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com'
  // });
}

// // ✅ Helper untuk encode UTF-8 + Base64 agar emoji di subject terbaca
// function encodeSubjectUTF8(subject) {
//   const utf8Bytes = Utilities.computeBytes(subject, Utilities.Charset.UTF_8);
//   const base64 = Utilities.base64Encode(utf8Bytes);
//   return `=?UTF-8?B?${base64}?=`;
// }
