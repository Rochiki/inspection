# Temuan Lapangan QSHE

Aplikasi milik **QSHE Department** untuk mencatat temuan inspeksi di lapangan dan menyusunnya menjadi laporan standar. Aplikasi ini:

- **Gratis.** Tidak memakai AI, akun berbayar, maupun server.
- **Jalan tanpa internet** setelah dibuka sekali, dan bisa dipasang di HP atau laptop seperti aplikasi biasa.
- **Merapikan otomatis.** Catatan singkat diubah jadi deskripsi temuan baku. Klausul ISO 9001:2026, ISO 14001:2026, ISO 45001:2018, kriteria SMK3 PP 50/2012, serta regulasi terkait dipilih otomatis dari kata kunci.
- **Punya alur review dua tahap.** Supervisor menyetujui dulu, lalu QSHE memberi persetujuan Final.
- **Membuat laporan PowerPoint dan Excel**, lalu menyimpannya ke folder (termasuk folder OneDrive laptop).

---

## 1. Struktur folder

| File / folder | Isi | Boleh diedit? |
|---|---|---|
| `index.html` | Kerangka halaman | Jarang |
| `css/style.css` | Tampilan. Warna diatur di bagian paling atas (`:root`) | Ya |
| `js/data-klausul.js` | Database klausul ISO, kriteria SMK3, dan regulasi | **Ya, paling sering** |
| `js/rapikan.js` | Kamus bahasa lapangan dan 24 topik kata kunci untuk "Rapikan tanpa AI" | **Ya, paling sering** |
| `js/app.js` | Logika aplikasi. Daftar area, PIC, dan supervisor bawaan ada di `DEFAULT_SETTINGS` (dekat baris 87) | Hati-hati |
| `js/data-ppt.js`, `img/` | Latar slide PowerPoint | Jarang |
| `sw.js` | Mode offline. **Angka `VERSI` wajib dinaikkan setiap update** | Hanya angka VERSI |
| `lib/`, `fonts/`, `licenses/` | Pustaka pihak ketiga, font, dan lisensinya | **Jangan diubah** |
| `manifest.webmanifest`, `icon-*.png` | Nama dan ikon saat dipasang di HP | Jarang |

---

## 2. Memasang di GitHub (sekali saja, ±10 menit)

1. Buat akun di **github.com** (gratis).
2. Klik **New repository**. Beri nama, misalnya `temuan-qshe`, pilih **Public**, lalu klik **Create repository**.
3. Klik **uploading an existing file**. Ekstrak zip ini, lalu seret **seluruh isi foldernya** ke halaman upload: file `index.html` dan kawan-kawan, plus folder `css`, `js`, `lib`, `fonts`, `img`, dan `licenses`. Jangan seret folder induknya. Klik **Commit changes**.
4. Buka **Settings → Pages**. Di bagian **Build and deployment**, pilih **Source: Deploy from a branch**, lalu **Branch: main**, folder **/ (root)**, dan klik **Save**.
5. Tunggu 1–2 menit, lalu muat ulang halaman Settings → Pages. Alamat aplikasi akan muncul, misalnya:
   `https://namaakun.github.io/temuan-qshe/`

> **Kenapa Public?** GitHub Pages gratis hanya tersedia untuk repository publik. Yang terlihat orang hanya **kode aplikasi**. **Data temuan tidak pernah tersimpan di GitHub**; data ada di HP/laptop masing-masing dan di folder OneDrive Anda. Hindari menaruh nama orang atau data sensitif di file kode (misalnya di `DEFAULT_SETTINGS`). Isi hal-hal itu lewat menu Pengaturan di aplikasi.

---

## 3. Berbagi akses

### a. Pengguna aplikasi (teknisi, supervisor, safety officer)

Mereka **tidak butuh akun GitHub**. Cukup kirim link aplikasinya, atau pakai tombol **Pengaturan → Bagikan link aplikasi**.

- **HP Android:** buka link di Chrome, tekan menu **⋮ → Instal aplikasi** (atau *Tambahkan ke layar utama*).
- **Laptop:** buka link di Chrome/Edge, lalu klik ikon **Instal** di ujung kanan kolom alamat.

### b. Tim pengembang (yang ikut mengedit aplikasi)

1. Setiap anggota membuat akun GitHub gratis.
2. Di repository, buka **Settings → Collaborators → Add people**, lalu masukkan username atau email mereka.
3. Mereka menerima undangan lewat email, lalu bisa ikut mengedit.

Kalau anggota tim makin banyak, buat **GitHub Organization** gratis (misalnya `qshe-dept`), lalu pindahkan repository ke sana lewat **Settings → Transfer**. Hak akses bisa diatur per tim. Perlu diingat, memindahkan repository **mengubah alamat aplikasi**; baca catatan penting di bagian 5.

---

## 4. Cara mengedit dan merilis update

1. Di halaman repository, tekan tombol **titik (`.`)** di keyboard. Editor **github.dev** akan terbuka di browser.
2. Edit file yang diperlukan, misalnya menambah kata kunci di `js/rapikan.js` atau klausul di `js/data-klausul.js`.
3. Buka `sw.js`, lalu naikkan versinya: `const VERSI = "v1";` menjadi `"v2"`, dan seterusnya.
4. Klik ikon **Source Control** di kiri, tulis keterangan perubahan, lalu klik **Commit & Push**.
5. Tunggu 1–2 menit. Pengguna mendapat versi baru setelah **membuka aplikasi dua kali**. Pembukaan pertama mengunduh update di latar belakang, dan pembukaan kedua memakainya.

**Saran:** buat repository kedua, misalnya `temuan-qshe-uji`, untuk mencoba perubahan. Setelah hasilnya benar, baru salin ke repository utama.

---

## 5. Catatan penting

- **Jangan mengganti nama akun GitHub atau memindahkan repository** setelah aplikasi dipakai tim. Data di HP terikat pada alamat `namaakun.github.io`. Kalau alamat berubah, aplikasi di HP akan tampil kosong. Kalau terpaksa pindah, minta semua orang mengirim **paket data** terlebih dahulu.
- **Folder bersama antar-laptop.** Beberapa laptop QSHE boleh menghubungkan aplikasinya ke **folder OneDrive/SharePoint yang sama** (Pengaturan → Folder penyimpanan). Data akan digabung otomatis: entri yang sama diambil versi terbarunya, dan entri yang dihapus tidak muncul lagi. Kalau OneDrive membuat salinan konflik (misalnya `database_temuan-NAMALAPTOP.json`), biarkan saja atau hapus. Data di aplikasi tetap aman.
- **HP Android tidak bisa terhubung ke folder.** Keterbatasan ini ada di browser, bukan di aplikasi. Data dari HP dikirim lewat **Unduh paket data (.json)**, lalu diimpor di laptop QSHE.
- **Menghapus data browser atau meng-uninstall aplikasi akan menghapus data yang belum dikirim atau disinkronkan.**

---

## 6. Membuat file APK (opsional)

1. Buka **pwabuilder.com**, tempel alamat aplikasi, lalu klik **Start**.
2. Pilih **Package for stores → Android → Generate**, lalu unduh zip-nya.
3. Bagikan file `.apk` ke HP tim. **Simpan file signing key baik-baik**, karena dibutuhkan untuk membuat versi APK berikutnya.

APK ini membuka alamat GitHub Pages yang sama, jadi setiap update di GitHub otomatis ikut. APK hanya perlu dibuat ulang kalau nama atau ikon aplikasi berubah.

---

## 7. Lisensi komponen

Aplikasi ini memakai komponen gratis berikut. Salinan lisensinya ada di folder `licenses/` dan **wajib tetap disertakan**.

| Komponen | Fungsi | Lisensi |
|---|---|---|
| ExcelJS 4.4.0 | Membuat file Excel | MIT |
| PptxGenJS 3.12.0 | Membuat file PowerPoint | MIT |
| Chart.js 4.4.1 | Grafik dashboard | MIT |
| Barlow & Barlow Condensed | Font | SIL Open Font License 1.1 |
