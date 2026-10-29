# Penghitung Penyusutan Peralatan & Mesin

Mengubah PDF **Daftar BMD pada Pengguna Barang – Peralatan & Mesin (FORMAT II.O.2.2)** dari SIPD e-BMD menjadi Excel
penyusutan (format kolom A–AL seperti `dpu_2024_NEW.xlsx`) dengan rumus Excel hidup, dengan format yang sama seperti contoh Anda (judul kolom 1–6 di B1:G1, freeze pane W4, AutoFilter), plus:
- sheet **Pivot 6 digit**: rekap per kode barang 6 digit (Nilai Perolehan, Akumulasi awal, Beban, Akumulasi Akhir, Nilai Buku) memakai rumus `SUMIFS` ke sheet `Sheet`, lengkap dengan Grand Total dan sel kontrol selisih terhadap total PDF (harus 0);
- sheet **1** (master masa manfaat) dan sheet **Temuan** (hasil pemeriksaan).

- PDF dibaca **di browser** (pdf.js); berkas tidak diunggah ke server.
- Master masa manfaat (bulan) dari tabel Supabase `masa_manfaat`; jika belum aktif, memakai `public/masa_manfaat.json`.
- Riwayat perhitungan per pengguna disimpan di Supabase (opsional, dengan login dan Row Level Security).

## Dua jenis laporan (tab)
- **Peralatan & Mesin** (FORMAT II.O.2.2): PDF saja; garis lurus per bulan dari nilai perolehan.
- **Gedung & Bangunan** (FORMAT II.O.2.3): PDF + **template saldo awal** `Format_penyusutan_gedung.xlsx` (kolom NIBAR, TGL PEROLEHAN, AKM akhir tahun lalu,
  Nilai buku akhir tahun lalu, Sisa masa manfaat (bln), Kapitalisasi tahun berjalan, Tanggal kapitalisasi, Tambahan masa manfaat).
  Data digabung berdasarkan **NIBAR (45 digit)**; hasilnya Excel berformat `diskop_penyusutan_2026.xlsx` (kolom A–AP, sheet Pivot, Temuan).
  Pastikan kolom NIBAR di template berformat **Teks** agar 45 digit tidak terpotong Excel.

## Metode
Garis lurus, tanpa nilai residu, per bulan: `beban/bulan = nilai perolehan ÷ masa manfaat (bulan)`.
Bulan terpakai = bulan sejak perolehan s.d. 31 Des (bulan perolehan dihitung penuh, bisa diganti "mulai bulan berikutnya"), dibatasi masa manfaat.
Beban tahun berjalan = (bulan terpakai tahun ini − tahun lalu) × beban/bulan.

## Pemeriksaan otomatis
Kode tidak ada di master · tanggal tidak valid/setelah periode · nilai ≠ harga × jumlah · NIBAR ganda · di bawah batas kapitalisasi (opsional) ·
**rekonsiliasi total tiap kelompok di PDF vs jumlah rincian** · info: aset habis disusutkan, tanggal 31 Desember (tanggal estimasi).

## Deploy
1. **Supabase**: buat project → SQL Editor → jalankan `supabase/schema.sql` → Table Editor → `masa_manfaat` → *Import CSV* `supabase/masa_manfaat_seed.csv`.
   Authentication → aktifkan Email. Ambil *Project URL* dan *anon key* (Settings → API).
2. **GitHub**: `git init && git add . && git commit -m "init" && git remote add origin <url> && git push -u origin main`.
3. **Vercel**: Add New Project → pilih repo → Framework *Vite* → Environment Variables `VITE_SUPABASE_URL` dan `VITE_SUPABASE_ANON_KEY` → Deploy.
4. Supabase → Authentication → URL Configuration: isi *Site URL* dengan domain Vercel.

Lokal: `npm install && cp .env.example .env && npm run dev`.
Metode Gedung: beban = nilai buku awal ÷ sisa masa manfaat × min(12, sisa); bila ada kapitalisasi, dipecah sebelum/sesudah bulan kapitalisasi (rumus sama dengan Excel acuan).
Uji akurasi (peralatan dan gedung): `DATA_DIR=/folder/berisi/dpu.pdf+dpu_2024_NEW.xlsx/ npm test` (membandingkan hasil dengan Excel acuan).
