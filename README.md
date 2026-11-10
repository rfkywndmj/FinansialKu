# FinansialKu - Jurnal Kas & Portofolio Pribadi

<p align="center">
  <img src="assets/icon-180.png" alt="Logo FinansialKu" width="108" height="108" style="border-radius: 24px;">
</p>

**FinansialKu** adalah aplikasi web pencatat keuangan pribadi dan portofolio reksa dana berbasis **Google Sheets + Google Apps Script**. Seluruh data tersimpan 100% di Google Sheet milikmu sendiri tanpa biaya server dan tanpa pihak ketiga.

---

## Fitur Utama

- **100% Data Milik Sendiri**: Berjalan di atas Google Sheets dan Google Apps Script pribadi dengan saldo awal `Rp 0` (kosong dan siap diisi sendiri).
- **Kalibrasi Saldo Instan**: Atur daftar bank/e-wallet sendiri dan sesuaikan saldo aktual kapan saja langsung dari aplikasi.
- **Auto-Sync Email Bank**: Membaca email notifikasi transaksi dari **Livin' by Mandiri** (QRIS, Pembayaran Tagihan, Transfer, Top-up E-Wallet) dan **Bebas by Bank MAS** secara otomatis setiap 10 menit.
- **Portofolio Reksa Dana**: Pantau modal, imbal hasil harian/bulanan, dan riwayat NAB reksa dana pasar uang.
- **Plafon Pengeluaran Harian (Rollover Budget)**: Batas harian otomatis menyesuaikan sisa anggaran hari sebelumnya.
- **Antarmuka Swiss Ledger**: Mendukung tema Terang & Gelap (Light/Dark Mode), pembaruan instan (Optimistic UI), serta mode layar penuh (**Standalone WebClip**) di iPhone dan Android.

---

## Struktur Folder

```text
FinansialKu/
├── Code.gs                    # Backend Google Apps Script (API, rumus Sheet, & parser email bank)
├── Index.html                 # Frontend aplikasi web (UI, grafik, tema terang/gelap)
├── appsscript.json            # Konfigurasi manifest Google Apps Script (Zona waktu Asia/Jakarta)
├── FinansialKu.mobileconfig   # Template profil iOS untuk instalasi layar penuh (Standalone) di iPhone
├── .clasp.json.example        # Contoh konfigurasi bagi pengguna CLI Clasp (opsional)
└── assets/                    # Ikon aplikasi (180x180, 192x192, 512x512)
```

---

## Panduan Pemakaian dari Awal sampai Akhir

Ikuti langkah berikut untuk memasang FinansialKu di akun Google kamu sendiri dari nol hingga siap dipakai di HP:

### Langkah 1: Buat Google Sheet Baru (Database)

1. Buka [Google Sheets](https://sheets.google.com) lalu klik **Blank spreadsheet** (Spreadsheet kosong).
2. Beri nama bebas pada bagian kiri atas, misalnya **`Database FinansialKu`**.
3. **Tidak perlu membuat tab atau kolom secara manual.** Aplikasi akan otomatis membuat tab `Pengeluaran`, `Pemasukan`, `Pengaturan`, dan `Investasi` beserta seluruh rumus `SUMIFS`-nya saat pertama kali dijalankan.

---

### Langkah 2: Pasang Kode di Google Apps Script

1. Di dalam Google Sheet yang baru kamu buat, klik menu bagian atas: **Extensions** (Ekstensi) > **Apps Script**.
2. Di halaman Apps Script yang terbuka, pastikan zona waktu berada di **GMT+7 (Asia/Jakarta)**:
   - Klik ikon roda gigi (**Project Settings**) di bilah kiri.
   - Pastikan **Time zone** diatur ke `(GMT+07:00) Western Indonesia Time - Jakarta` (atau centang *Show "appsscript.json" manifest file in editor* lalu salin isi `appsscript.json`).
3. Kembali ke menu **Editor (`< >`)** di bilah kiri:
   - **File `Code.gs`**:
     1. Klik file `Code.gs` yang sudah ada, hapus seluruh isinya, lalu tempel (paste) seluruh isi dari file [`Code.gs`](./Code.gs) di repositori ini.
     2. *(Opsional)* Jika kamu membuka Apps Script langsung dari menu **Extensions > Apps Script** di dalam Google Sheet, variabel `SHEET_ID` di baris ke-3 **boleh dibiarkan kosong (`''`)**. Jika kamu membuat script terpisah di luar Sheet, isi `SHEET_ID` dengan ID dari URL Google Sheet kamu.
     3. *(Opsional - untuk fitur Auto-Sync Email Bank)* Isi variabel `OWNER_NAME` di baris ke-5 dengan nama lengkapmu sesuai yang tertera di rekening bank (contoh: `'BUDI SANTOSO'`). Ini dipakai untuk mengenali apakah transfer ditujukan ke rekeningmu sendiri di bank lain.
   - **File `Index.html`**:
     1. Klik tombol **`+` (Add a file)** di sebelah tulisan **Files** > pilih **HTML**.
     2. Beri nama persis **`Index`** (huruf `I` kapital, tanpa menulis `.html` karena ditambahkan otomatis).
     3. Hapus kode bawaannya, lalu tempel seluruh isi dari file [`Index.html`](./Index.html) di repositori ini.
4. Tekan **Ctrl + S** (Windows) atau **Cmd + S** (Mac) untuk menyimpan kedua file tersebut.

---

### Langkah 3: Berikan Izin Akses (Otorisasi Pertama)

1. Di bagian atas editor Apps Script, ubah pilihan fungsi di sebelah tombol **Debug** menjadi **`setupSheets`**, lalu klik tombol **Run**.
2. Akan muncul jendela **Authorization required**:
   - Klik **Review permissions** lalu pilih akun Google kamu.
   - Jika muncul peringatan *"Google hasn't verified this app"*, klik **Advanced** (Lanjutan) di kiri bawah, lalu klik **Go to Untitled project (unsafe)** / **Lanjutkan**. *(Peringatan ini normal karena script dibuat oleh kamu sendiri di akun pribadimu)*.
   - Klik **Allow** (Izinkan).
3. Tunggu hingga proses *Execution completed* selesai. Sekarang coba lihat Google Sheet kamu: tab `Pengeluaran`, `Pemasukan`, `Pengaturan`, dan `Investasi` sudah otomatis terbuat dengan saldo Rp 0!

---

### Langkah 4: Deploy sebagai Web App

1. Di pojok kanan atas editor Apps Script, klik tombol biru **Deploy** > pilih **New deployment**.
2. Klik ikon roda gigi di sebelah *Select type* > pilih **Web app**.
3. Isi pengaturan berikut:
   - **Description**: `FinansialKu`
   - **Execute as**: `Me` (Email kamu)
   - **Who has access**:
     - Pilih **`Anyone`** jika kamu ingin memasang aplikasi ini di layar utama iPhone (mode layar penuh / Standalone WebClip) agar tidak mental ke halaman login Google.
     - Pilih **`Only myself`** jika kamu hanya membukanya lewat browser biasa yang sudah login akun Google.
4. Klik **Deploy**.
5. Salin **Web app URL** yang muncul (berakhiran `/exec`). Buka URL tersebut di browser untuk mencoba aplikasimu!

---

### Langkah 5: Atur Rekening & Isi Saldo Awal Sendiri

Saat pertama kali dibuka, seluruh saldo di aplikasi adalah **Rp 0** (kosong). Untuk menyesuaikannya dengan uangmu saat ini:

1. **Sesuaikan Daftar Bank & E-Wallet**:
   - Buka tab **Pengaturan** (ikon paling kanan di navigasi bawah).
   - Ketuk menu **Daftar Rekening & Dompet**.
   - Hapus bank yang tidak kamu gunakan (klik ikon tempat sampah), atau ketik nama rekening/dompet baru (misal: `BCA`, `Mandiri`, `Bank Jago`, `SeaBank`, `GoPay`, `ShopeePay`, `Tunai`) lalu klik **Tambah**.
2. **Atur Batas Harian & Kategori**:
   - Masih di tab **Pengaturan**, sesuaikan **Batas Pengeluaran Harian** (misal `50.000` per hari) dan kategori pemasukan/pengeluaran sesuai kebutuhanmu.
3. **Isi Saldo Awal Masing-Masing Rekening**:
   - Kembali ke tab **Beranda**.
   - Pada bagian **02 Saldo per Rekening**, ketuk **ikon pensil** di pojok kanan atas kartu rekening yang ingin diisi.
   - Masukkan nominal saldo asli yang ada di rekening/dompet kamu saat ini pada kolom **Saldo Aktual**, lalu ketuk **Simpan Penyesuaian**.
   - Lakukan untuk semua rekeningmu. Total kas akan langsung terhitung otomatis dan akurat!

---

### Langkah 6: Pasang Aplikasi di Layar Utama HP (Full Screen)

#### Pengguna Android (Chrome)
1. Buka **Web app URL** kamu di aplikasi Google Chrome.
2. Ketuk ikon **titik tiga** di pojok kanan atas browser.
3. Pilih **Install app** atau **Tambahkan ke Layar Utama (Add to Home screen)**.

#### Pengguna iPhone / iOS (Mode Layar Penuh Tanpa Bar Safari + Logo Custom)
Karena Google Apps Script berjalan di dalam *iframe*, fitur *Add to Home Screen* bawaan Safari tidak menampilkan ikon custom. Gunakan profil WebClip resmi berikut agar tampil layar penuh dengan logo emerald:

1. Buka aplikasi **FinansialKu** kamu di laptop/Mac, masuk ke tab **Pengaturan** > **Pasang Aplikasi di HP (PWA)**.
2. Klik tombol **Unduh Profil Standalone iPhone (`.mobileconfig`)**. File `FinansialKu.mobileconfig` yang sudah berisi URL Web App milikmu akan terunduh otomatis.
   *(Atau kamu juga bisa membuka file `FinansialKu.mobileconfig` di repo ini menggunakan teks editor, lalu ganti tulisan `https://script.google.com/macros/s/GANTI_DENGAN_DEPLOYMENT_ID_KAMU/exec` dengan Web App URL milikmu).*
3. **Kirim file ke iPhone**:
   - **Cara AirDrop (Mac ke iPhone)**: Klik kanan file `FinansialKu.mobileconfig` di Finder Mac > pilih **Share** > **AirDrop** > pilih iPhone kamu.
   - **Cara Email (Windows/Linux/Mac)**: Kirim file `FinansialKu.mobileconfig` sebagai lampiran email ke alamat emailmu, lalu buka email tersebut melalui aplikasi **Apple Mail** bawaan di iPhone dan ketuk lampirannya.
4. Di layar iPhone akan muncul notifikasi **Profile Downloaded** (Profil Diunduh).
5. Buka aplikasi **Settings (Pengaturan)** di iPhone > ketuk menu **Profile Downloaded** di bagian paling atas > ketuk **Install** di pojok kanan atas > masukkan PIN iPhone > ketuk **Install** lagi.
6. Selesai! Aplikasi **FinansialKu** akan muncul di Home Screen iPhone dengan logo emerald dan terbuka layar penuh tanpa bar Safari.

---

### Langkah 7: Sinkronisasi Otomatis Email Bank (Opsional)

FinansialKu dapat membaca email bukti transaksi yang masuk ke Gmail kamu secara otomatis setiap 10 menit:
- **Bank yang didukung bawaan**:
  - **Livin' by Mandiri**: Pembayaran QRIS, Pembayaran Tagihan/Merchant (Traveloka, dll.), Transfer Rupiah, dan Top-up E-Wallet (otomatis memisahkan pokok transfer ke akun tujuan seperti GoPay dan mencatat biaya admin secara terpisah).
  - **Bebas by Bank MAS**: Pembayaran QRIS dan Transfer.
- **Keamanan Saldo**:
  - Hanya transaksi yang terjadi **sejak tanggal inisialisasi buku kas** (`sync_start_date` di tab `Pengaturan`) yang akan dimasukkan, sehingga email transaksi masa lalu tidak akan merusak saldo awal yang baru kamu atur.
- **Kategorisasi Otomatis**:
  - Aturan pencocokan nama merchant ke kategori dapat kamu ubah sesuka hati pada variabel `MERCHANT_RULES` di bagian atas file `Code.gs`.

---

## Kredit & Atribusi

Dirancang dan dikembangkan oleh **Polarize__**.
