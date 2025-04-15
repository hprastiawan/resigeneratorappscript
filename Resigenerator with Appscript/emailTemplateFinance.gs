function sendEmailToFinance(trxData, blobs) {
  const {
    idTransaksi, nama, tipe, program,
    tanggal, jumlah, status, channel
  } = trxData;

  const subjectRaw = `[Finance Notification] ${idTransaksi} - ${nama} - ${tipe} ${program} - ${status} 🔔`;

  const bannerUrl = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Email%20Header%20Banner%20-%20Phincon%20Academy/Email%20Banner%20Phincon%20Academy%202025.png';
  const logoFooter = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/Phincon%20Academy%20-%20Logo%20Footer.png';
  const linkedinIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/linkedin.png';
  const instagramIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/instagram.png';
  const whatsappIcon = 'https://raw.githubusercontent.com/hprastiawan/emailbanner/refs/heads/main/whatsapp.png';

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; background-color: #f9f9f9; padding: 30px; max-width: 640px; margin: auto; border-radius: 10px;">
      <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="border-radius: 8px; padding: 20px;">
        <tr>
          <td>
            <img src="${bannerUrl}" alt="Phincon Academy Email Banner" style="width: 100%; max-width: 600px; border-radius: 8px;" />
          </td>
        </tr>
        <tr>
          <td style="padding-top: 20px;">
            <h2 style="text-align: center; color: #333;">Notifikasi Dana Masuk</h2>
            <p>Hi Finance Team,</p>
            <p>Berikut adalah detail pembayaran dari peserta:</p>
            <table width="100%" cellpadding="8" cellspacing="0" style="font-size: 14px; border-collapse: collapse;">
              <tr style="background-color: #f3fdf6;">
                <td><strong>ID Transaksi</strong></td>
                <td><strong>${idTransaksi}</strong></td>
              </tr>
              <tr>
                <td>Nama</td>
                <td>${nama}</td>
              </tr>
              <tr>
                <td>Program</td>
                <td>${tipe} ${program}</td>
              </tr>
              <tr>
                <td>Tanggal Transaksi</td>
                <td>${tanggal}</td>
              </tr>
              <tr>
                <td>Jumlah</td>
                <td><strong>${formatRupiah(jumlah)}</strong></td>
              </tr>
              <tr>
                <td>Channel Pembayaran</td>
                <td>${channel}</td>
              </tr>
              <tr>
                <td>Status Pembayaran</td>
                <td><strong>${status}</strong></td>
              </tr>
            </table>
            <p style="margin-top: 16px;">Bukti pembayaran dan file resi terlampir dalam email ini.</p>
            <p>Terima kasih 🙏</p>
          </td>
        </tr>
        <tr>
          <td style="padding: 30px; background-color: #f3f3f3; text-align: center; border-radius: 8px;">
            <img src="${logoFooter}" alt="Phincon Academy Logo Footer" style="height: 40px; margin-bottom: 10px;" />
            <div style="font-size: 13px; color: #555; line-height: 1.4;">
              Gandaria 8 Office Tower, 8th Floor<br>
              Jl. Arteri Pd. Indah No.10, RT.9/RW.6 Kby. Lama Utara,<br>
              Kec. Kby. Lama, Kota Jakarta Selatan<br>
              Jakarta 12240, Indonesia
            </div>
            <div style="margin-top: 10px;">
              <a href="https://www.linkedin.com/school/phincon-academy/" target="_blank" style="display: inline-block;">
                <img src="${linkedinIcon}" alt="LinkedIn" style="height: 24px; margin: 0 6px;" />
              </a>
              <a href="https://www.instagram.com/phinconacademy?igsh=MzRlODBiNWFlZA%3D%3D" target="_blank" style="display: inline-block;">
                <img src="${instagramIcon}" alt="Instagram" style="height: 24px; margin: 0 6px;" />
              </a>
              <a href="https://api.whatsapp.com/send/?phone=6281119970372&text&type=phone_number&app_absent=0" target="_blank" style="display: inline-block;">
                <img src="${whatsappIcon}" alt="WhatsApp" style="height: 24px; margin: 0 6px;" />
              </a>
            </div>
          </td>
        </tr>
      </table>
    </div>
  `;

  //Versi Prod Baru
  MailApp.sendEmail({
    to: "gloria.daeli@phintraco.com, anisa.sangraningrum@phincon.com, maria.permatasari@phintraco.com",
    subject: subjectRaw,
    htmlBody: htmlBody,
    attachments: blobs,
    name: 'Phincon Academy System',
    cc: "academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com, phinconacademy@gmail.com",
    bcc: "phinconacademy@gmail.com"
    //cc: "hendra.prastiawan4@gmail.com"
  });

  //Versi Dev schema Baru
  // MailApp.sendEmail({
  //   to: "hendra.prastiawan2@gmail.com",
  //   subject: subjectRaw,
  //   htmlBody: htmlBody,
  //   attachments: blobs,
  //   name: 'Phincon Academy System',
  //   cc: "hendra.prastiawan4@gmail.com"
  // });

  //Versi Prod Lama - Works
  // GmailApp.sendEmail("gloria.daeli@phintraco.com, anisa.sangraningrum@phincon.com, maria.permatasari@phintraco.com", encodeSubjectUTF8(subjectRaw), '', {
  //   htmlBody: htmlBody,
  //   attachments: blobs,
  //   name: 'Phincon Academy System',
  //   cc: "academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com"
  //   //cc: "hendra.prastiawan4@gmail.com"
  // });

  // Untuk Kebutuhan Testing schema lama (pake GAS Service)
  //   GmailApp.sendEmail("hendra.prastiawan2@gmail.com", encodeSubjectUTF8(subjectRaw), '', {
  //     htmlBody: htmlBody,
  //     attachments: blobs,
  //     name: 'Phincon Academy System',
  //     cc: "academy@phincon.com, payment.academy@phincon.com, ghama.bayu@phincon.com, tasya.jannah@phincon.com"
  //     // cc: "hendra.prastiawan4@gmail.com"
  //   });
}

function formatRupiah(nominal) {
  if (!nominal || nominal === "-") return "-";
  const number = parseInt(nominal.toString().replace(/[^\d]/g, ''));
  if (isNaN(number)) return "-";
  return `Rp ${number.toLocaleString('id-ID')},-`;
}
