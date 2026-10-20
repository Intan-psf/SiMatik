# SiMatik — Sistem Asesmen Matematika Terintegrasi dan Kredibel

Aplikasi ini memakai satu Google Spreadsheet sebagai database. Siswa mengerjakan ujian lewat browser HP; guru mengelola siswa, ujian, soal, dan penilaian melalui spreadsheet.

Setelah guru mengisi atau mengubah nilai ISIAN/ESSAY pada kolom kuning, trigger edit spreadsheet memperbarui nilai guru ke detail dan menghitung kembali nilai akumulasi. Jalankan **📝 Ujian → Siapkan spreadsheet** sekali setelah memperbarui Apps Script untuk mengaktifkan trigger edit berizin.

## Berkas

- `Code.gs` — fungsi server dan menu guru.
- `index.html` — tampilan web siswa, menjadi halaman utama Vercel.
- `appsscript.json` — konfigurasi Apps Script.
- `api/relay.js` — jembatan API Vercel ke backend Apps Script.
- `api.php` dan `config.php` — jembatan PHP ke backend Apps Script untuk penggunaan lokal.
- `Template_Ujian.xlsx` — template impor dengan contoh data dan petunjuk kolom.
- `generate_template.py` — skrip untuk membuat ulang template Excel.

## 1. Siapkan spreadsheet dan Apps Script

1. Buat Google Spreadsheet kosong, lalu buka **Ekstensi → Apps Script**.
2. Ganti isi file `Code.gs` dengan isi berkas `Code.gs` pada folder ini.
3. Klik **+ → HTML**, beri nama persis `Index`, lalu salin isi `index.html`.
4. Buka **Setelan project** (ikon roda gigi), aktifkan **Tampilkan file manifes "appsscript.json" di editor**, lalu ganti isinya dengan berkas `appsscript.json` di folder ini.
5. Simpan project dan kembali ke spreadsheet. Muat ulang tab spreadsheet.
6. Dari menu **📝 Ujian**, pilih **Siapkan spreadsheet**. Saat diminta izin Google, tinjau dan berikan izin.

Menu akan membuat sheet `Siswa`, `Daftar Ujian`, dan struktur contoh `KUIS1` (sheet soal, peserta, jawaban ringkas satu baris per siswa, serta detail jawaban). Data contoh hanya ditambahkan bila sheet terkait masih belum berisi data, sehingga persiapan ulang tidak menimpa isi sheet yang sudah ada.
Jika kolom rekap menampilkan `#ERROR!`, pastikan `Code.gs` sudah diperbarui lalu jalankan kembali **📝 Ujian → Siapkan spreadsheet**. Rumus rekap akan dibuat ulang mengikuti lokal spreadsheet, termasuk lokal Indonesia.

Contoh token ujian adalah `MTK8`. Ubah atau hapus contoh siswa/ujian setelah mencoba aplikasi.

## 2. Isi siswa dan soal

- Isi sheet `Siswa` dengan kolom **Kelas**, **No Absen**, dan **Nama**.
- Atur ujian pada `Daftar Ujian`. Kolom **Kode** harus sama dengan akhiran nama sheet `Soal_`, `Peserta_`, dan `Jawaban_`. Token ujian peka huruf besar/kecil: siswa harus mengetik token persis sama dengan isi kolom **Token**.
- Masukkan soal ke sheet `Soal_<KODE>`. Tipe yang tersedia: `PG`, `ISIAN`, dan `ESSAY`.
- Untuk PG, isi opsi dan kunci dengan huruf opsi spreadsheet, misalnya `B`. Opsi yang kosong tidak ditampilkan.
- Untuk ISIAN, kunci jawaban bersifat opsional sebagai referensi guru; penilaiannya dilakukan manual seperti ESSAY.
- Isi bobot tiap soal. PG dinilai otomatis berdasarkan kunci. Guru memberi nilai ISIAN dan ESSAY pada kolom kuning `Nilai Guru <nomor soal>` di `Jawaban_<KODE>`, dari 0 sampai bobot soal.
- `Jawaban_<KODE>` menampilkan satu baris per siswa, dengan jawaban pada kolom `Jawaban <nomor soal>` dan kolom nilai guru di sebelah jawaban ISIAN/ESSAY. Tepat setelah `Skor PG` hanya ada satu kolom rekap, `Nilai Akumulasi (/100)`, yang menampilkan nilai seluruh skor PG dan nilai guru dibanding total bobot soal.
- Kolom `Nilai Akumulasi (/100)` memakai rumus lembar kerja yang langsung merujuk ke `Skor PG`, kolom nilai guru, dan bobot tiap soal. Jadi nilainya berubah seketika saat guru mengisi atau mengubah skor manual; tidak perlu mengetik rumus sendiri.
- Nilai guru dibatasi sesuai bobot pada `Soal_<KODE>`. Contoh, jika bobot soal 2, nilai yang diterima hanya 0–2; masukan di atas bobot akan ditolak dan diberi peringatan.
- Setiap siswa hanya dapat mengumpulkan ujian satu kali. Jika sudah selesai lalu mencoba masuk lagi, aplikasi menampilkan pesan bahwa ujian sudah dikerjakan. Guru dapat memakai menu **Reset siswa** bila perlu mengizinkan percobaan ulang.
- `Detail_<KODE>` menyimpan rincian setiap jawaban per soal untuk pemeriksaan dan pembahasan. Nilai guru dijumlahkan dari kolom kuning di `Jawaban_<KODE>`, disalin ke detail, dan diperbarui otomatis pada akumulasi nilai.
- Jika spreadsheet lama masih memiliki format `Jawaban_<KODE>` satu baris per soal, simpan kode Apps Script yang baru lalu jalankan kembali **Siapkan spreadsheet**. Sheet lama akan dipindah ke `Detail_<KODE>`, dan sheet `Jawaban_<KODE>` baru akan dibuat tanpa menghapus jawaban lama.
- Link Google Drive untuk gambar harus dapat dilihat oleh siswa yang membuka web app.
- Teks soal, opsi, dan pembahasan dapat memakai LaTeX di antara `$...$` atau `$$...$$`.

Gunakan menu **📝 Ujian → Buat ujian baru** untuk menambahkan ujian beserta sheet-nya. Ujian baru berstatus `Nonaktif`; periksa soal dan token, lalu aktifkan melalui kolom **Status** ketika siap.

## 3. Deploy web app

1. Di editor Apps Script, pilih **Deploy → New deployment**.
2. Klik ikon roda gigi dan pilih **Web app**.
3. Atur **Execute as** menjadi **Me** (pemilik/deployer) dan **Who has access** menjadi **Anyone**. Manifest menggunakan `ANYONE_ANONYMOUS`, sehingga siapa pun yang memiliki tautan dapat membuka halaman masuk.
4. Klik **Deploy**, berikan izin yang diminta, lalu salin **Web app URL** yang berakhir dengan `/exec`.
5. Kirim tautan dan token ujian kepada siswa. Simpan spreadsheet agar hanya guru yang dapat mengaksesnya; siswa tidak memerlukan akses ke spreadsheet.

Jika menu akses publik tidak tersedia di akun Google Workspace, admin domain mungkin membatasi deployment anonim.

## 4. Menjalankan halaman melalui PHP lokal

Pilihan ini menjalankan **halaman siswa** melalui PHP, tetapi data tetap dibaca/ditulis di Google Spreadsheet oleh backend Apps Script. PHP tidak menggantikan database atau Apps Script.

1. Pastikan spreadsheet sudah disiapkan dan project Apps Script sudah terhubung dengannya.
2. Salin isi `Code.gs` terbaru ke editor Apps Script, simpan, lalu pilih **Deploy → Manage deployments**.
3. Klik ikon pensil pada deployment Web app, pilih **New version**, lalu klik **Deploy**. Pastikan aksesnya tetap **Anyone** dan eksekusi sebagai pemilik. `doPost` diperlukan sebagai API untuk PHP.
4. Buka `config.php`, lalu ganti `PASTE_DEPLOYMENT_ID` pada URL contoh dengan **Web app URL** deployment yang berakhir dengan `/exec`.
5. Pastikan PHP 8.1+ dan ekstensi **cURL** tersedia. Periksa dengan:

   ```powershell
   php -v
   php -m | findstr /I curl
   ```

6. Buka PowerShell di folder aplikasi dan jalankan server lokal:

   ```powershell
   php -S 0.0.0.0:8000
   ```

7. Di komputer yang menjalankan PHP, buka `http://localhost:8000/index.html`.

Jangan membuka `index.html` dengan klik ganda (`file://`), karena halaman perlu memanggil endpoint PHP `api.php`. Jangan juga menggunakan server statis yang tidak menjalankan PHP.

### Membuka dari HP di jaringan Wi-Fi yang sama

`localhost` hanya menunjuk ke perangkat itu sendiri. Agar HP membuka halaman dari server PHP di komputer, komputer dan HP harus berada di jaringan Wi-Fi yang sama:

1. Jalankan server dengan `php -S 0.0.0.0:8000` seperti di atas.
2. Di komputer, jalankan `ipconfig` dan cari **IPv4 Address** adaptor Wi-Fi, misalnya `192.168.1.20`.
3. Buka `http://192.168.1.20:8000/index.html` di HP. Jika Windows meminta izin firewall, izinkan PHP pada jaringan privat.

Alamat lokal ini hanya berfungsi saat komputer menyala dan HP berada di jaringan yang sama. Untuk membagikan ujian melalui internet, gunakan URL deployment Apps Script, Vercel, atau tempatkan aplikasi PHP di hosting HTTPS yang dapat diakses siswa. Jangan bagikan spreadsheet dengan izin edit publik.

## 5. Deploy ke Vercel

1. Pastikan `Code.gs` sudah disimpan dan deployment Apps Script aktif dengan akses **Anyone** serta URL berakhiran `/exec`.
2. Push file project ke GitHub, lalu di Vercel pilih **Add New → Project** dan impor repository SiMatik.
3. Biarkan **Framework Preset** sebagai **Other** dan **Root Directory** sebagai folder project. Tidak perlu Build Command atau Output Directory.
4. Di pengaturan project Vercel, buka **Settings → Environment Variables** dan tambahkan `APPS_SCRIPT_WEB_APP_URL` dengan nilai URL Web App Apps Script yang berakhiran `/exec`. Pilih environment **Production** (tambahkan **Preview** juga jika ingin mengetes preview), lalu simpan. Jangan menambahkan tanda kutip.
5. Klik **Deploy**. Setelah selesai, buka domain Vercel; halaman akan dimuat dari `index.html`.
6. Jika URL deployment Apps Script berubah, perbarui environment variable tersebut di Vercel lalu lakukan **Redeploy**.

## 6. Buat QR code untuk siswa

Gunakan menu berbagi QR di browser Chrome (jika tersedia), atau buka layanan pembuat QR code yang dipercaya. Masukkan URL web app `/exec`, unduh gambar QR, lalu cetak atau tampilkan di kelas. Uji hasil pindai dengan HP dalam mode privat sebelum dibagikan. QR hanya berisi tautan; token ujian tetap sebaiknya dibagikan terpisah.

## 7. Mengelola ujian

- **Reset siswa** menghapus catatan peserta dan jawaban siswa terpilih agar dapat mengulang. Gunakan setelah memastikan identitas dan kode ujiannya benar.
- **Arsipkan ujian** menyalin sheet soal, peserta, dan jawaban sebagai nilai tetap ke spreadsheet `Arsip Ujian`, lalu menghapus sheet ujian tersebut dari spreadsheet aktif dan mengubah status ujian menjadi `Diarsipkan`.
- **Bersihkan cache** diperlukan bila perubahan data tidak segera terlihat. Perubahan manual pada sheet `Siswa`, `Daftar Ujian`, dan `Soal_*` juga menghapus cache terkait secara otomatis.

## Catatan penggunaan

- Timer mengikuti waktu server dan batas waktu ujian; jawaban sementara juga disimpan pada perangkat dan disinkronkan berkala saat tersedia koneksi.
- Sesi aktif baru menggantikan sesi perangkat sebelumnya. Keluar dari aplikasi/tab dan keluar dari layar penuh dapat dicatat sebagai pelanggaran, sesuai kemampuan browser dan perangkat.
- Fullscreen dan deteksi perpindahan aplikasi tidak dapat dijamin oleh semua browser, terutama iPhone. Siswa tetap harus mengikuti aturan ujian dan guru sebaiknya mengawasi ujian.
- Jawaban kunci dan pembahasan tidak dikirim saat siswa mengerjakan; penilaian otomatis dilakukan di server. Setelah ujian selesai, pembahasan hanya ditampilkan bila pengaturannya `Ya`.
- Untuk pengaturan nilai `Setelah ditutup`, pastikan kolom **Tutup** diisi. Jika kosong, nilai tidak otomatis ditampilkan karena belum ada waktu tutup yang ditentukan.
- Jangan ubah atau membagikan kolom tersembunyi `Sesi`, `Jawaban Sementara`, dan `Data` di sheet peserta.

## Pemeriksaan pertama yang disarankan

1. Buka tautan web app sebagai siswa contoh `8A`, absen `1`, token `MTK8`; pastikan konfirmasi nama dan aturan tampil benar.
2. Kerjakan soal, pindah soal, refresh untuk mencoba pemulihan sesi, dan pastikan jawaban tersimpan.
3. Kumpulkan ujian dan periksa `Nilai Akumulasi (/100)` setelah kolom `Skor PG` pada `Jawaban_KUIS1`, serta nilai total pada sheet `Peserta_KUIS1`.
4. Uji satu pelanggaran (misalnya pindah tab) dan pastikan jeda serta hitungannya tercatat di sheet peserta.
5. Coba isi nilai guru pada kolom kuning `Nilai Guru 4` atau `Nilai Guru 6` di `Jawaban_KUIS1`; nilai di atas bobot soal harus ditolak. Pastikan `Nilai Akumulasi (/100)` ikut berubah.

Untuk membuat ulang file Excel setelah mengubah generator, jalankan:

```powershell
python -m pip install openpyxl
python .\generate_template.py
```
