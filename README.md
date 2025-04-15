# Resi & Email Automation System 📩

Sistem otomatisasi untuk pembuatan **Resi Pembayaran** dan pengiriman **Email Konfirmasi** peserta, dibangun menggunakan **Google Apps Script**.

---

## 🎯 Fitur Utama

- 🆔 **Generate ID Registrasi & Transaksi** otomatis berdasarkan tipe program, lokasi, dan tanggal transaksi.
- 📄 **Generate PDF Resi** berbasis template Google Slides.
- 📬 **Kirim Email Otomatis** ke peserta dengan lampiran resi pembayaran.
- 📤 **Upload Bukti Transfer** langsung ke Google Drive per baris data.
- 💰 **Notifikasi ke Tim Finance** dengan email grup + attachment (resi + bukti transfer).
- 🔐 **Proteksi Kolom & Header** otomatis untuk menjaga integritas data.
- 🔄 **Menu Refresh Halaman** untuk update dropdown, validasi, dan status file.

---

## 📁 Struktur Sheet & Folder

- `Form Responses 1` → Data utama peserta & pembayaran.
- `Data Email ke Finance` → Sinkronisasi status upload & kirim email ke Finance.
- `Kirim ke Finance` → Queue untuk validasi dan pengiriman batch.
- 📂 Google Drive Folder:
  - `Resi PDF` (`FOLDER_ID_RESI`)
  - `Bukti Transfer` (`FOLDER_ID_TRANSFER`)

---

## 🛠️ Teknologi

- **Google Apps Script** (GAS)
- **Google Sheets**
- **Google Slides (Template)**
- **Google Drive API**
- HTML Sidebar (`uploadBuktiTransferWeb.html`) untuk upload file

---

## ⚙️ Menu Otomatis

Tersedia dalam spreadsheet:

```
🏢 HOME
├── 🔄 Refresh Halaman
├── 🆔 Generate ID
│   └── 🔢 Generate ID Transaksi & Registrasi
├── 🧾 Buat Resi
│   ├── 📕 Untuk Baris Ini
│   ├── 📘📗 Untuk Baris Terpilih
│   └── 🚚 Seluruh Data
├── 📧 Kirim Resi ke Email Peserta
│   ├── 👤 Baris Ini
│   ├── 👥 Baris Terpilih
│   └── 👥👥 Seluruh Data
└── 📩 Email ke Finance
    ├── 🧲 Muat Ulang Data
    ├── 📤 Upload Bukti Bayar
    └── 🚀 Kirim Email
```

---

## ✅ Status & Keamanan

- Validasi data wajib: Email, ID Transaksi, ID Registrasi
- Status dikontrol di kolom:
  - `Status Resi PDF`
  - `File dalam Folder`
  - `Send Email Status`
  - `Status Bukti Transfer`
  - `Status Kirim ke Finance`

---

## 🚀 Kontribusi

Project ini digunakan untuk mengotomasi administrasi pembayaran peserta. Kontribusi terbuka untuk pengembangan lanjutan seperti integrasi WhatsApp API atau dashboard analytics.

---

## 📃 Lisensi

Proyek ini bersifat internal. Untuk penggunaan atau kontribusi luar, harap hubungi maintainer terlebih dahulu.
