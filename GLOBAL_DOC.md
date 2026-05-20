# GLOBAL: Panduan Teknis Proyek Excel Dasar

Tujuan dokumen ini: menyatukan aturan teknis, layout sheet, cara menghasilkan dataset, dan fitur "copy data" agar GitHub Copilot atau kontributor lain dapat membaca spesifikasi teknis secara konsisten. Dokumen ini bersifat global — file `week*.md` dan `midproject.md` cukup berisi deskripsi materi dan contoh data.

## instruksi teknis
markdown dokumen digunakan untuk menyusun file latihan .html, contoh week3.md untuk menciptakan file week3.html, dan seterusnya

### teknis html:
- layout menggunakan tailwind dan font google
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Latihan Pertemuan 1: Excel Basic</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    body { font-family: 'Inter', sans-serif; }
    .excel-header { background-color: #217346; }
    .code-block { background-color: #f5f5f5; border-left: 4px solid #217346; }
</style>

- responsive design
- lihat week7.html sebagai referensi


## Ringkasan File-File Markdown
- `midproject.md` — Deskripsi kasus mid project, data input contoh, dan langkah solusi. Lihat [midproject.md](midproject.md).
- `week3.md` — Mathematical & Text Functions. Lihat [week3.md](week3.md).
- `week4.md` — Date-Time Functions. Lihat [week4.md](week4.md).
- `week5.md` — Conditional Logic (IF). Lihat [week5.md](week5.md).
- `week6.md` — Lookup Functions (VLOOKUP/HLOOKUP/INDEX-MATCH). Lihat [week6.md](week6.md).
- `week7.md` — Latihan Lookup terstruktur dan sample dataset. Lihat [week7.md](week7.md).
- `week8.md` - Latihan rumus financial (PMT, PV, NPER, RATE, IPMT, PPMT), Lihat [week8.md]

Jika ada file `.md` baru, tambahkan ringkasan satu baris di bagian atas file ini.

## Aturan Umum untuk Semua `week*.md` dan `midproject.md`
- Hanya jelaskan konteks pembelajaran, contoh data, dan tugas praktikum.
- Jangan masukkan instruksi teknis implementasi (layout sheet, macro, Power Query). Semua instruksi teknis berada di `GLOBAL_DOC.md`.
- Sertakan metadata singkat di awal file dalam format kunci: `Kode Modul`, `Pertemuan`, `Durasi`, `Prasyarat`.

Contoh metadata header di `weekX.md`:

Kode Modul: EXCEL-00X
Pertemuan: X
Durasi Praktik: 90 menit
Prasyarat: Week 1..X-1

## Layout Workbook (Konvensi Sheet)
- Sheet wajib: `DATA_INPUT`, `PROSES`, `DASHBOARD`.
- `DATA_INPUT`: hanya berisi raw data yang di-copy dari soal (tidak ada formula kompleks).
- `PROSES`: semua formula dan kolom bantu berada di sini. Formula ditulis pada baris pertama data (baris 2) lalu di-copy ke bawah.
- `DASHBOARD`: ringkasan KPI, grafik, dan insight.

### Header Kolom Standar di Sheet `PROSES` (Kasus Penjualan)
- A: `Invoice ID`
- B: `Sales Person`
- C: `Kode Produk`
- D: `Region`
- E: `Channel`
- F: `Tanggal Transaksi`
- G: `Qty`
- H: `Nama Produk` (lookup)
- I: `Harga Unit` (lookup)
- J: `Kategori` (lookup)
- K: `Total Sales` (=G * I)
- L: `Validasi` (OK / Perlu Koreksi)
- M: `Prefix Invoice` (LEFT)
- N: `Nomor Urut Invoice` (RIGHT)
- O: `Hari` (DAY)
- P: `Bulan` (MONTH)
- Q: `Tahun` (YEAR)
- R: `Umur Transaksi` (DATEDIF)

Sesuaikan header ini untuk kasus selain penjualan; jaga konsistensi nama kolom agar dashboard dan template formula tetap bekerja.

## Template Formula (gunakan nama sheet persis)
- `Nama Produk`: `=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,2,FALSE),"Kode Tidak Valid")`
- `Harga Unit`: `=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,3,FALSE),0)`
- `Kategori`: `=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,4,FALSE),"Unknown")`
- `Total Sales`: `=G2*I2`
- `Validasi`: `=IF(H2="Kode Tidak Valid","Perlu Koreksi","OK")`
- Ambil target bulan (contoh HLOOKUP di `DATA_INPUT`): `=HLOOKUP("Mar",DATA_INPUT!$T$1:$Z$3,2,FALSE)`

Gunakan referensi absolut (`$A$1:$C$10`) untuk rentang lookup agar formula aman saat dicopy.

## Generate Dataset — Panduan & Template
1. Format file sumber: CSV atau Excel (`.xlsx`).
2. Kolom minimal untuk transaksi: `Invoice ID, Sales Person, Kode Produk, Region, Channel, Tanggal Transaksi, Qty`.
3. Kolom master produk minimal: `Kode Produk, Nama Produk, Harga Unit, Kategori` (letakkan di `DATA_INPUT` mulai kolom `M` s.d. `P`).
4. Contoh CSV header transaksi:

Invoice ID,Sales Person,Kode Produk,Region,Channel,Tanggal Transaksi,Qty

5. Untuk dataset uji, sediakan satu baris invalid key (mis: `PRD-999`) agar latihan validasi berfungsi.

### Skrip singkat (opsional): buat CSV contoh
Simpan contoh berikut sebagai `dataset_transactions.csv` untuk import:

Invoice ID,Sales Person,Kode Produk,Region,Channel,Tanggal Transaksi,Qty
INV-0326-001,Budi Santoso,PRD-001,Malang,Online,2026-03-01,2
INV-0326-002,Ani Wijaya,PRD-004,Surabaya,Offline,2026-03-03,4
INV-0326-003,Rudi Firmansah,PRD-003,Malang,Online,2026-03-05,3
INV-0326-004,Budi Santoso,PRD-002,Blitar,Online,2026-03-08,1
INV-0326-005,Ani Wijaya,PRD-008,Kediri,Offline,2026-03-10,6
INV-0326-006,Rudi Firmansah,PRD-006,Malang,Online,2026-03-15,2
INV-0326-007,Budi Santoso,PRD-005,Surabaya,Online,2026-03-20,5
INV-0326-008,Ani Wijaya,PRD-999,Malang,Offline,2026-03-25,2

## Fitur "Copy Data" — Praktik yang Direkomendasikan
- Manual copy/paste: Copy A:G dari `DATA_INPUT` ke `PROSES` (gunakan Paste Values jika data berasal dari sheet lain).
- Power Query (disarankan untuk dataset besar):
  - `Data` → `Get Data` → `From File` → `From Workbook` / `From CSV` lalu load ke `PROSES` sebagai query.
  - Keuntungan: mudah refresh, transformasi kolom, trimming spasi dan cast tipe data.
- Macro VBA (opsional, auto copy):

  Sub CopyDataToProses()
    Sheets("DATA_INPUT").Range("A2:G100").Copy
    Sheets("PROSES").Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
  End Sub

  - Sesuaikan range (`A2:G100`) sesuai panjang data. Simpan macro di workbook macro-enabled (`.xlsm`).

## Quality Rules & Best Practices
- Gunakan `TRIM()` pada kolom kunci sebelum lookup untuk menghindari spasi tersembunyi: `=TRIM(C2)`.
- Pastikan kolom kunci (`Kode Produk`) unik pada master data.
- Simpan lookup table (master) di `DATA_INPUT` di kolom `M:P` agar semua formula di `PROSES` konsisten.
- Gunakan format tanggal standar (yyyy-mm-dd) untuk menghindari ambiguity saat parsing.

## Instruksi untuk GitHub Copilot (atau agen otomatis)
1. Jika diminta membuat `weekX.md`, fokuskan hanya pada konten pengajaran dan contoh data. Jangan menulis instruksi teknis layout atau macro.
2. Jika diminta menghasilkan file Excel contoh, ikuti header dan formula template di dokumen ini.
3. Saat membuat kode (VBA/Power Query), tempatkan contoh di `scripts/` atau lampirkan ke `GLOBAL_DOC.md`.

## Checklist saat menyerahkan tugas
- `DATA_INPUT` berisi raw data yang sama dengan soal.
- `PROSES` berisi formula di baris pertama data dan sudah dicopy ke semua baris.
- `DASHBOARD` berisi KPI minimal (Total, Average, Min, Max, Count, Valid/Invalid counts).
- File Excel tidak menyembunyikan formula.
- Jika menggunakan macro, lampirkan `.xlsm` dan dokumentasikan macro singkat.
