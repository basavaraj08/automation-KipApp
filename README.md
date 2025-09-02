KipApp.py — Notepad / README
==================================================
Automasi input KiPapp (https://webapps.bps.go.id/kipapp/#/pelaksanaan-aksi) berbasis Selenium
yang mengambil data dari Google Sheets lalu mengisi form secara otomatis.

Fungsi Utama
--------------------------------------------------
1) Ambil data dari Google Sheets (worksheet: "dashboard").
2) Buka halaman KiPapp dan pilih periode SKP (dropdown "Pilih SKP").
3) Untuk tiap baris data:
   - Pilih "Rencana Kinerja SKP" dengan fuzzy matching (mirip ≥ 60%).
   - Klik Add → buka modal.
   - Centang "Gunakan periode tanggal" & "Masukkan ke Capaian SKP".
   - Pilih tanggal awal–akhir dari datepicker (berdasarkan kolom E & F).
   - Isi "Deskripsi Kegiatan", "Deskripsi Capaian", dan "Link Data Dukung".
   - Klik Save.

Persyaratan (Dependencies)
--------------------------------------------------
- Python 3.10+
- Google Sheets API:
  - gspread
  - google-auth (google.oauth2.service_account)
- Selenium 4+
- Chrome & ChromeDriver (versi harus kompatibel)

Install cepat (virtualenv opsional):
  pip install gspread google-auth google-auth-oauthlib selenium

Kredensial Google Sheets
--------------------------------------------------
1) Buat Service Account & unduh credentials.json.
2) Bagikan Google Sheet ke email Service Account (Viewer/Editor).
3) Simpan file credentials.json di direktori yang sama dengan KipApp.py
   atau ubah path di kode:
     creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)

Sumber Data (Google Sheets)
--------------------------------------------------
- URL contoh di kode:
  https://docs.google.com/spreadsheets/d/19dP73kPkAvKLqYbOlaWYevdtsBH4dcAPPovhW2Fp7SM/edit?usp=sharing
- Worksheet: "dashboard"
- Data yang diambil: range A:J (baris pertama dianggap header).

Pemetaan Kolom (berdasarkan kode)
--------------------------------------------------
A (col 0) → deskripsi_kegiatan
B (col 1) → rencana_kinerja (teks untuk dropdown "Rencana Kinerja SKP")
E (col 4) → tanggal_awal  (format diharapkan dd/mm/yyyy atau mm/dd/yyyy → kode ambil DD via split("/")[1])
F (col 5) → tanggal_akhir (format sama dengan kolom E)
G (col 6) → deskripsi_capaian
H (col 7) → link_data_dukung
(C dan D tidak digunakan pada kode saat ini; I dan J juga belum digunakan)

Catatan tanggal:
- Kode mengekstrak nilai DD dari string tanggal dengan:
    start_date_dd = str(int(tanggal_awal.split("/")[1]))
    end_date_dd   = str(int(tanggal_akhir.split("/")[1]))
  Pastikan urutan pemisah tanggal sesuai data (jika format berbeda, sesuaikan indeks).

Konfigurasi Selenium / Chrome
--------------------------------------------------
- ChromeOptions:
  --start-maximized, --log-level=3
- *Opsional* gunakan profil Chrome agar tetap login SSO:
    options.add_argument("user-data-dir=C:/Users/<USER>/AppData/Local/Google/Chrome/User Data")
    options.add_argument("profile-directory=Profile 2")
- Path ChromeDriver (ubah sesuai lokasi Anda):
    chrome_driver_path = "D:/WISNU/_dev/automation/chromedriver.exe"

Langkah Menjalankan
--------------------------------------------------
1) Pastikan credentials.json & paket Python sudah siap.
2) Pastikan ChromeDriver cocok dengan versi Chrome.
3) (Opsional) Set profil Chrome agar sesi login KiPapp tersimpan.
4) Jalankan:
     python KipApp.py
5) Script akan:
   - Membuka KiPapp (zoom 80%).
   - Menunggu ±30 detik untuk halaman siap/login manual jika perlu.
   - Menutup modal notifikasi (OK) jika ada.
   - Memilih periode SKP (lihat bagian "Ubah Periode Bulan").
   - Iterasi tiap baris dan submit.

Ubah Periode Bulan (Dropdown "Pilih SKP")
--------------------------------------------------
- Locator saat ini:
  //li[contains(text(), "1 Januari - 31 Desember (Bulan Juli)")]
- Ubah teks di atas sesuai opsi di aplikasi:
  contoh: "1 Januari - 31 Desember (Bulan Agustus)"
- Blok terkait di kode:
    bulan_option = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, '//li[contains(text(), "1 Januari - 31 Desember (Bulan Juli)")]'))
    )

Fuzzy Matching "Rencana Kinerja SKP"
--------------------------------------------------
- Normalisasi teks: lowercase + hapus spasi berlebih.
- Cari kemiripan dengan SequenceMatcher; ambang (threshold) 0.6 (60%).
- Jika cocok persis → pilih langsung.
- Jika tidak, pilih opsi dengan kemiripan tertinggi ≥ 60%.
- Jika < 60% → baris dilewati & lanjut ke baris berikutnya.

Interaksi Form
--------------------------------------------------
- Checkbox:
  - "Gunakan periode tanggal" (label)
  - "Masukkan ke Capaian SKP" (ID: form-add_isCapaianSKP) → dicek, jika belum checked
- Tanggal:
  - Klik input tanggal (ID: form-add_tanggal) → buka datepicker
  - Pilih tanggal awal & akhir berdasarkan DD yang diekstrak
  - Menutup datepicker dengan klik ke body
- Input teks (dengan verifikasi & retry sekali):
  - form-add_kegiatan      ← deskripsi_kegiatan (A)
  - form-add_capaian       ← deskripsi_capaian (G)
  - form-add_dataDukung    ← link_data_dukung (H)

Penanganan Error & Logging
--------------------------------------------------
- Jika dropdown atau elemen tidak ditemukan → cetak peringatan & lanjut/skip.
- Simpan screenshot pada kegagalan input field / tanggal:
  - error_<nama_field>.png
  - error_general.png
- Progress dicetak ke console:
  "Processing row i/n", "Entry completed", dsb.

Tips Keandalan
--------------------------------------------------
- Perbesar WebDriverWait (dari 5 → 10/15 detik) bila jaringan lambat.
- Gunakan driver.execute_script('arguments[0].click()', el) untuk bypass overlay.
- Pastikan format tanggal konsisten; jika tidak, parsing perlu disesuaikan.
- Pertimbangkan locator yang lebih stabil (data-testid/role, bila ada).
- Gunakan try/except di tiap langkah kritikal agar satu baris gagal tidak menghentikan batch.

Checklist Cepat
--------------------------------------------------
[ ] credentials.json ada & Sheet di-share ke service account
[ ] URL spreadsheet & worksheet sesuai
[ ] Versi Chrome = ChromeDriver match
[ ] (Opsional) user-data-dir & profile-directory diaktifkan
[ ] XPath bulan SKP sudah sesuai periode yang diinginkan
[ ] Format tanggal di Sheet sesuai parsing (split("/")[1] = DD)
[ ] Uji 1–2 baris dulu sebelum full batch

Lisensi & Keamanan
--------------------------------------------------
- Jangan commit credentials.json ke repo publik.
- Masukkan credentials.json ke .gitignore.
- Kelola token/API key melalui variabel lingkungan (.env) jika diperlukan.

Catatan Tambahan
--------------------------------------------------
- Kode memanggil print total records dua kali (duplikasi aman, bisa dirapikan).
- Anda dapat menambahkan logging ke file & opsi resume untuk baris yang gagal.
- Untuk headless mode:
    options.add_argument("--headless=new")
  (perhatikan elemen visual/overlay mungkin berbeda di headless).

