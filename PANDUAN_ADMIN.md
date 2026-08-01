# Panduan Fitur Admin — Ruang Ujian Tryout TKA

## Apa yang ditambahkan

Sebelumnya, untuk menyetujui pendaftar / melihat hasil ujian, guru harus buka Google Sheets langsung.
Sekarang ada **menu Admin di dalam web ujian itu sendiri** (`index.html`), jadi guru cukup buka
link ujian → klik "Masuk sebagai Admin" di footer → masukkan password.

Panel Admin bisa:
- **Tab Pendaftaran**: lihat semua siswa yang mendaftar, Setujui / Tolak / Reset Kode / Hapus per baris,
  atau "Setujui Semua yang Menunggu" sekali klik.
- **Tab Hasil Ujian**: lihat semua hasil tryout (skor, benar/salah/kosong) + rata-rata skor, dan Hapus data.
- Statistik ringkas: jumlah menunggu, disetujui, jumlah hasil, rata-rata skor.

Tidak perlu buka Google Sheets lagi untuk kerja harian — Sheets tetap ada sebagai backup data mentah.

## Cara pasang (2 langkah)

### 1. Update Apps Script
1. Buka Google Sheet ujian Anda → **Extensions/Ekstensi → Apps Script**.
2. Hapus isi script lama, ganti dengan isi file **`apps_script_v6.js`** (lampiran ini).
3. Klik **Deploy → Manage deployments → Edit (pensil) → Deploy** lagi (pakai deployment/URL yang sama,
   supaya `SHEET_WEBHOOK_URL` di `index.html` tidak perlu diubah).
4. Kembali ke Google Sheet, **refresh halaman** — menu "Ruang Ujian" akan muncul lagi (dengan 1 item baru).
5. Klik menu **Ruang Ujian → Atur Password Admin** dan set password yang Anda mau.
   (Kalau tidak diset, password default sementara adalah `ubah-password-ini` — sebaiknya segera diganti.)

### 2. Update file web ujian
Ganti `index.html` yang di-hosting (GitHub Pages Anda) dengan file **`index.html`** (lampiran ini).
Tidak ada konfigurasi tambahan yang perlu diubah — `SHEET_WEBHOOK_URL` yang sudah ada tetap dipakai.

## Cara pakai
1. Buka link ujian seperti biasa.
2. Scroll ke footer paling bawah → klik **"Masuk sebagai Admin"**.
3. Masukkan password admin yang sudah diset di langkah 1.5.
4. Kelola pendaftaran & hasil dari panel yang muncul.

## Catatan keamanan
- Password admin dikirim sebagai bagian dari URL (pola yang sama dengan fitur login siswa yang sudah ada
  sebelumnya). Ini cukup untuk penggunaan internal kelas/sekolah, tapi bukan tingkat keamanan enterprise.
  Jangan bagikan password admin ke siswa, dan ganti password ini kapan saja lewat menu Sheets.
- Aksi "Hapus" bersifat permanen (baris dihapus dari Sheet), tidak ada tempat sampah/undo.
