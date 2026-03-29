# Latihan Pertemuan 7: Lookup Functions Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-007
- **Pertemuan**: 7
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Menengah Awal (Beginner-Intermediate)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 s.d. Week 6

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi lookup dalam konteks bisnis nyata
   - Menggunakan VLOOKUP() untuk pencarian data vertikal
   - Menggunakan HLOOKUP() untuk pencarian data horizontal
   - Menerapkan exact match dan approximate match sesuai kebutuhan
   - Mengombinasikan lookup dengan IFERROR untuk validasi hasil

2. **C4 (Analyze)** - Menganalisis data master dan transaksi dengan lookup
   - Mengintegrasikan data transaksi dengan master data produk
   - Menganalisis konsistensi kode, harga, dan kategori
   - Menyusun ringkasan insight berdasarkan hasil lookup
   - Merumuskan rekomendasi perbaikan kualitas data

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: VLOOKUP() - Vertical Lookup

#### 1.1 VLOOKUP() Dasar (Exact Match)
**Definisi**: Mencari nilai pada kolom pertama tabel referensi secara vertikal, lalu mengambil nilai dari kolom tertentu pada baris yang sama.

**Sintaks**:
```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Parameter**:
- lookup_value: Nilai kunci yang dicari (misal: kode produk)
- table_array: Tabel referensi (master data)
- col_index_num: Nomor kolom yang akan diambil dari table_array
- range_lookup: FALSE untuk exact match, TRUE untuk approximate match

**Return Value**: Nilai dari kolom target pada baris yang sesuai.

**Kegunaan Bisnis**:
- Mengambil nama produk dari kode produk
- Mengambil harga standar dari master produk
- Mengambil kategori produk untuk pelaporan
- Validasi transaksi terhadap data master

**Contoh**:
```
CV Komputer Indonesia - Lookup Nama Produk
- Tujuan: Ambil Nama Produk berdasarkan Kode Produk di transaksi
- Formula: =VLOOKUP(B2, $M$2:$P$9, 2, FALSE)
- Hasil: Nama produk sesuai kode

Interpretasi: Transaksi otomatis terhubung ke master produk tanpa input manual berulang
```

**Catatan Praktis**:
- Gunakan FALSE jika ingin hasil tepat 100% (exact match).
- Gunakan absolute reference (misal $M$2:$P$9) saat copy formula.

---

#### 1.2 VLOOKUP() Multi-Kolom (Harga dan Kategori)
**Contoh Formula**:
```
Ambil Harga Unit:
=VLOOKUP(B2, $M$2:$P$9, 3, FALSE)

Ambil Kategori:
=VLOOKUP(B2, $M$2:$P$9, 4, FALSE)
```

**Interpretasi**:
Satu lookup key (Kode Produk) bisa dipakai untuk mengambil beberapa atribut penting sekaligus.

---

#### 1.3 VLOOKUP() Approximate Match
**Definisi**: Mencari nilai terdekat jika lookup_value tidak harus sama persis.

**Contoh Kasus**:
Menentukan diskon berdasarkan tabel tier penjualan.

**Contoh Formula**:
```
=VLOOKUP(G2, $T$2:$U$6, 2, TRUE)
```

**Catatan**:
- Tabel referensi harus diurutkan naik.
- Cocok untuk tier, grading, atau skala komisi.

---

### Blok 2: HLOOKUP() - Horizontal Lookup

#### 2.1 HLOOKUP() Dasar
**Definisi**: Mencari nilai pada baris pertama tabel referensi secara horizontal, lalu mengambil nilai dari baris tertentu pada kolom yang sama.

**Sintaks**:
```
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

**Parameter**:
- lookup_value: Nilai kunci yang dicari di baris pertama
- table_array: Tabel horizontal
- row_index_num: Nomor baris yang diambil hasilnya
- range_lookup: FALSE (exact) atau TRUE (approximate)

**Return Value**: Nilai pada baris target sesuai kolom hasil pencarian.

**Kegunaan Bisnis**:
- Mengambil target bulanan dari tabel target horizontal
- Mengambil persentase bonus berdasarkan bulan
- Membaca parameter bisnis yang disusun per kolom periode

**Contoh**:
```
CV Komputer Indonesia - Lookup Target Bulanan
- Lookup: bulan "Mar"
- Formula: =HLOOKUP("Mar", $B$12:$M$14, 2, FALSE)
- Hasil: Target sales bulan Maret

Interpretasi: Memudahkan penarikan target berdasarkan header bulan
```

---

#### 2.2 HLOOKUP() untuk Multi-Baris Parameter
**Contoh Formula**:
```
Target Sales:
=HLOOKUP("Apr", $B$12:$M$14, 2, FALSE)

Target Margin:
=HLOOKUP("Apr", $B$12:$M$14, 3, FALSE)
```

**Interpretasi**:
Satu bulan dapat memetakan beberapa parameter target (sales, margin, dll).

---

### Blok 3: Integrasi Lookup + IFERROR

#### 3.1 Menangani Kode Tidak Ditemukan
**Definisi**: IFERROR digunakan agar formula lookup tidak menampilkan error mentah.

**Contoh Formula**:
```
=IFERROR(VLOOKUP(B2, $M$2:$P$9, 2, FALSE), "Kode Tidak Valid")
```

**Kegunaan Bisnis**:
- Menandai data transaksi yang tidak terdaftar di master
- Memudahkan data cleansing dan quality control

**Interpretasi**:
Jika kode tidak ada di master, hasil tidak #N/A tetapi pesan informatif.

---

#### 3.2 Validasi Lookup Harga
**Contoh Formula**:
```
=IFERROR(VLOOKUP(B2, $M$2:$P$9, 3, FALSE), 0)
```

**Interpretasi**:
Jika harga tidak ditemukan, nilai default 0 dipakai agar perhitungan lanjutan tetap berjalan.

---

### Blok 4: Best Practice Lookup untuk Data Bisnis

#### 4.1 Prinsip Penting
- Pastikan kolom kunci unik (tidak duplikat untuk exact match).
- Gunakan FALSE untuk transaksi harian agar akurat.
- Gunakan TRUE hanya saat tabel referensi tersortir dan memang butuh nearest match.
- Gunakan IFERROR untuk output yang lebih bersih.
- Pisahkan master data dan data transaksi secara jelas.

#### 4.2 Kesalahan Umum
- Salah col_index_num/row_index_num.
- Table range bergeser karena tidak absolute reference.
- Mencari key yang ternyata ada spasi tersembunyi.
- Menggunakan approximate match pada data yang butuh exact match.

---

### Blok 5 (Opsional): INDEX + MATCH - Belajar Mandiri

**Status Materi**:
- Materi ini **opsional** untuk eksplorasi mandiri di luar kelas/perkuliahan.
- Tidak masuk checklist penilaian utama pada praktikum kelas.

#### 5.1 MATCH() Dasar
**Definisi**: Mengembalikan posisi relatif suatu nilai dalam range.

**Sintaks**:
```
=MATCH(lookup_value, lookup_array, [match_type])
```

**Contoh**:
```
=MATCH("PRD-003", $M$2:$M$9, 0)
```
Interpretasi: Menemukan baris ke-berapa kode PRD-003 di master.

#### 5.2 INDEX() Dasar
**Definisi**: Mengembalikan nilai dari range berdasarkan nomor baris/kolom.

**Sintaks**:
```
=INDEX(array, row_num, [column_num])
```

**Contoh**:
```
=INDEX($N$2:$N$9, 3)
```
Interpretasi: Mengambil nilai pada posisi baris ke-3.

#### 5.3 Kombinasi INDEX + MATCH
**Contoh Formula**:
```
=INDEX($N$2:$N$9, MATCH(B2, $M$2:$M$9, 0))
```

**Interpretasi**:
Mengambil nama produk berdasarkan kode produk dengan cara yang lebih fleksibel dibanding VLOOKUP.

**Kelebihan INDEX-MATCH (untuk eksplorasi mandiri)**:
- Bisa lookup ke kiri maupun kanan.
- Tidak bergantung nomor kolom statis.
- Lebih tahan terhadap perubahan struktur kolom master.

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: CV Komputer Indonesia - Transaction + Master Product

**Konteks Bisnis**:
Tim operasional menerima data transaksi harian berisi kode produk. Untuk analisis dan pelaporan, data transaksi harus dilengkapi nama produk, kategori, dan harga dari master data.

### Struktur Data Transaksi

| No | Kolom | Tipe Data | Keterangan | Contoh |
|----|-------|-----------|-----------|---------|
| A | No Transaksi | Text | ID transaksi | TRX-0701 |
| B | Kode Produk | Text | Kunci lookup | PRD-001 |
| C | Region | Text | Wilayah transaksi | Malang |
| D | Channel | Text | Sumber order | Online |
| E | Qty | Integer | Jumlah unit | 2 |
| F | Nama Produk | Text | Hasil lookup | Laptop Office 14 |
| G | Harga Unit | Currency | Hasil lookup | 8,500,000 |
| H | Kategori | Text | Hasil lookup | Laptop |
| I | Total Sales | Currency | Qty x Harga Unit | 17,000,000 |
| J | Validasi | Text | Status lookup | OK / Kode Tidak Valid |

### Struktur Master Data Produk

| Kolom | Isi |
|------|-----|
| M | Kode Produk |
| N | Nama Produk |
| O | Harga Unit |
| P | Kategori |

### Sample Data Transaksi (8 Baris)

```
No Transaksi | Kode Produk | Region   | Channel | Qty
TRX-0701     | PRD-001     | Malang   | Online  | 2
TRX-0702     | PRD-004     | Surabaya | Offline | 4
TRX-0703     | PRD-003     | Malang   | Online  | 3
TRX-0704     | PRD-002     | Blitar   | Online  | 1
TRX-0705     | PRD-008     | Kediri   | Offline | 6
TRX-0706     | PRD-006     | Malang   | Online  | 2
TRX-0707     | PRD-005     | Surabaya | Online  | 5
TRX-0708     | PRD-999     | Malang   | Offline | 2
```

### Sample Master Data Produk (M2:P9)

```
Kode Produk | Nama Produk       | Harga Unit | Kategori
PRD-001     | Laptop Office 14  | 8500000    | Laptop
PRD-002     | Printer Inkjet X  | 2200000    | Printer
PRD-003     | Monitor 24" FHD   | 3200000    | Monitor
PRD-004     | Laptop Pro 15     | 12500000   | Laptop
PRD-005     | Keyboard Mechanical| 950000    | Aksesoris
PRD-006     | Printer Laser Pro | 4800000    | Printer
PRD-007     | Mouse Wireless    | 250000     | Aksesoris
PRD-008     | UPS 1200VA        | 1750000    | Power
```

**Catatan Data**:
- Baris TRX-0708 sengaja menggunakan kode PRD-999 untuk simulasi invalid lookup.
- Mahasiswa diminta menangani kasus invalid dengan IFERROR.

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Lookup
*Fokus: VLOOKUP, HLOOKUP, IFERROR - Durasi: 40 menit*

#### Latihan 1.1: Lengkapi Data Transaksi dengan VLOOKUP
**Waktu**: 15 menit

**Instruksi**:
1. Input data transaksi (A:E) dan master data (M:P).
2. Isi kolom F (Nama Produk):
   ```
   =IFERROR(VLOOKUP(B2, $M$2:$P$9, 2, FALSE), "Kode Tidak Valid")
   ```
3. Isi kolom G (Harga Unit):
   ```
   =IFERROR(VLOOKUP(B2, $M$2:$P$9, 3, FALSE), 0)
   ```
4. Isi kolom H (Kategori):
   ```
   =IFERROR(VLOOKUP(B2, $M$2:$P$9, 4, FALSE), "Unknown")
   ```

**Analisis**:
- Berapa transaksi dengan kode valid?
- Kode mana yang invalid?

---

#### Latihan 1.2: Hitung Total Sales dan Validasi
**Waktu**: 12 menit

**Instruksi**:
1. Isi kolom I (Total Sales):
   ```
   =E2 * G2
   ```
2. Isi kolom J (Validasi):
   ```
   =IF(F2="Kode Tidak Valid", "Perlu Koreksi", "OK")
   ```
3. Copy formula ke semua baris transaksi.

**Analisis**:
- Berapa total sales dari transaksi valid saja?
- Dampak transaksi invalid terhadap laporan harian?

---

#### Latihan 1.3: HLOOKUP untuk Target Bulanan
**Waktu**: 13 menit

**Instruksi**:
1. Buat tabel horizontal target bulanan (misal Jan-Dec) di area berbeda.
2. Ambil target bulan tertentu (misal Mar) dengan HLOOKUP:
   ```
   =HLOOKUP("Mar", $B$20:$M$22, 2, FALSE)
   ```
3. Ambil parameter lain (misal target margin) pada baris berikutnya:
   ```
   =HLOOKUP("Mar", $B$20:$M$22, 3, FALSE)
   ```

**Analisis**:
- Apakah total sales valid sudah memenuhi target Maret?
- Jika belum, berapa gap terhadap target?

---

### LEVEL 2: Analisis Lanjut & Dashboard Lookup
*Fokus: Integrasi Lookup + Ringkasan Bisnis - Durasi: 50 menit*

#### Latihan 2.1: Dashboard Kualitas Data Lookup
**Waktu**: 15 menit

**Instruksi**:
Buat panel metrik berikut:
```
- Total Transaksi: =COUNTA(A2:A9)
- Transaksi Valid: =COUNTIF(J2:J9, "OK")
- Transaksi Invalid: =COUNTIF(J2:J9, "Perlu Koreksi")
- % Valid Data: =COUNTIF(J2:J9, "OK")/COUNTA(A2:A9)
```

Tambahkan status KPI:
```
=IF([%Valid] >= 0.95, "ON TRACK", "PERLU PERBAIKAN")
```

**Analisis**:
- Bagaimana kualitas data transaksi saat ini?
- Apa prioritas perbaikan data master/kode input?

---

#### Latihan 2.2: Dashboard Sales per Kategori
**Waktu**: 15 menit

**Instruksi**:
Buat ringkasan sales berbasis kategori hasil lookup:
```
- Total Sales Laptop: =SUMIF(H2:H9, "Laptop", I2:I9)
- Total Sales Printer: =SUMIF(H2:H9, "Printer", I2:I9)
- Total Sales Monitor: =SUMIF(H2:H9, "Monitor", I2:I9)
- Total Sales Aksesoris: =SUMIF(H2:H9, "Aksesoris", I2:I9)
```

Tambahkan kategori dominan:
```
Gunakan MAX + IF bertingkat atau analisis manual berbasis nilai terbesar.
```

**Analisis**:
- Kategori paling dominan terhadap revenue?
- Kategori mana yang perlu didorong promosi?

---

#### Latihan 2.3: Insight dan Rekomendasi Operasional
**Waktu**: 20 menit

**Instruksi**:
1. Tulis minimal 3 insight berbasis hasil lookup.
2. Tulis minimal 3 rekomendasi operasional.
3. Setiap insight harus menyertakan angka pendukung dari dashboard.

**Contoh Format Insight**:
```
Insight 1: Tingkat validasi data sebesar [x%], masih di bawah target 95%,
sehingga perlu standarisasi input kode produk.

Insight 2: Kategori [nama kategori] menyumbang revenue tertinggi sebesar [angka],
layak diprioritaskan untuk strategi upselling.
```

---

## 📘 LATIHAN MANDIRI (OPSIONAL - DI LUAR KELAS)

### Topik: INDEX + MATCH

**Status**:
- Tidak masuk checklist penilaian utama kelas.
- Direkomendasikan untuk latihan transisi dari VLOOKUP ke metode lookup yang lebih fleksibel.

### Tugas Mandiri 1: Lookup Nama Produk dengan INDEX-MATCH

Gunakan formula:
```
=INDEX($N$2:$N$9, MATCH(B2, $M$2:$M$9, 0))
```

Tambahkan IFERROR:
```
=IFERROR(INDEX($N$2:$N$9, MATCH(B2, $M$2:$M$9, 0)), "Kode Tidak Valid")
```

### Tugas Mandiri 2: Lookup Harga Unit dengan INDEX-MATCH

Gunakan formula:
```
=IFERROR(INDEX($O$2:$O$9, MATCH(B2, $M$2:$M$9, 0)), 0)
```

### Tugas Mandiri 3: Refleksi Perbandingan

Bandingkan VLOOKUP vs INDEX-MATCH dalam 3 poin:
1. Fleksibilitas struktur tabel
2. Kemudahan maintenance formula
3. Risiko error saat kolom master berubah

---

## 📋 INSTRUKSI PENGUMPULAN

### Format File
- Nama file: **NAMA_NIM_Week7_LookupFunctions.xlsx**
- Sheet 1: Data transaksi + formula
- Sheet 2: Dashboard ringkasan (recommended)
- Semua formula visible (tidak di-hide)

### Checklist Pengerjaan (90 menit):
- [ ] LEVEL 1 (40 min): Latihan 1.1, 1.2, 1.3 selesai tanpa error
- [ ] LEVEL 2 (50 min): Latihan 2.1, 2.2, 2.3 selesai + insight
- [ ] Dashboard menampilkan metrik validitas data dan sales kategori
- [ ] File save sesuai format nama

### Catatan Checklist:
- Latihan INDEX-MATCH bersifat mandiri (opsional) dan tidak dinilai dalam checklist utama kelas.

### Rubric Penilaian:

| Aspek | Kriteria | Bobot |
|-------|----------|-------|
| **C3 (Apply)** | Formula VLOOKUP, HLOOKUP, IFERROR, IF diinput benar; hasil lookup akurat | 50% |
| **C4 (Analyze)** | Dashboard validasi data dan insight operasional jelas, berbasis angka, dan relevan | 50% |

**Passing**: Score >= 70 (C3 >= 35 + C4 >= 35)

---

## 💡 TIPS & TROUBLESHOOTING

### Common Mistakes:

1. **Col Index Salah di VLOOKUP**:
   ```
   ❌ SALAH: =VLOOKUP(B2, $M$2:$P$9, 5, FALSE)
   ✅ BENAR: =VLOOKUP(B2, $M$2:$P$9, 2/3/4, FALSE)
   ```

2. **Lupa Absolute Reference**:
   ```
   ❌ SALAH: =VLOOKUP(B2, M2:P9, 2, FALSE)
   ✅ BENAR: =VLOOKUP(B2, $M$2:$P$9, 2, FALSE)
   ```

3. **Exact Match vs Approximate Match Tertukar**:
   ```
   ❌ SALAH (untuk kode produk): =VLOOKUP(B2, $M$2:$P$9, 2, TRUE)
   ✅ BENAR: =VLOOKUP(B2, $M$2:$P$9, 2, FALSE)
   ```

4. **#N/A Tidak Ditangani**:
   - Gunakan IFERROR agar output ramah pengguna.

5. **HLOOKUP Salah Baris Output**:
   - Pastikan row_index_num sesuai baris parameter yang ingin diambil.

### Troubleshooting Tips:
- Jika hasil #N/A: cek key tidak ditemukan atau ada spasi tersembunyi.
- Jika hasil tidak sesuai: cek col_index/row_index.
- Jika formula berubah saat drag: cek absolute reference.
- Jika approximate match aneh: pastikan tabel referensi diurutkan naik.

### Function Alternatives:
- Lookup vertikal sederhana -> gunakan VLOOKUP.
- Lookup horizontal parameter -> gunakan HLOOKUP.
- Lookup fleksibel -> gunakan INDEX-MATCH (latihan mandiri).

---

## 📖 INSTRUKSI UNTUK AGENT / TEACHING ASSISTANT

### Learning Objectives (Detailed):

1. **Lookup Mastery**:
   - Mahasiswa memahami kapan memakai VLOOKUP vs HLOOKUP.
   - Mahasiswa mampu menggunakan exact match untuk data transaksi.
   - Mahasiswa mampu menambahkan IFERROR untuk validasi data.

2. **Data Quality Awareness**:
   - Mahasiswa dapat membedakan transaksi valid vs invalid lookup.
   - Mahasiswa dapat mengaitkan kualitas data dengan akurasi laporan.

3. **Business Interpretation**:
   - Mahasiswa mampu menyusun insight berbasis hasil lookup.
   - Mahasiswa mampu memberi rekomendasi operasional dari data.

### Common Misconceptions to Clarify:

| Misconception | Correction |
|---|---|
| VLOOKUP otomatis selalu tepat | Harus pakai FALSE untuk exact match data transaksi |
| #N/A berarti formula rusak total | Bisa jadi lookup key memang tidak ada di master |
| HLOOKUP sama persis dengan VLOOKUP | Prinsip sama, orientasi tabel berbeda (horizontal vs vertikal) |
| INDEX-MATCH wajib menggantikan VLOOKUP | VLOOKUP tetap valid untuk kasus sederhana |
| IFERROR tidak penting | Sangat penting untuk kualitas output dan dashboard |

### Assessment Tips:
- Cek akurasi hasil lookup terhadap master data secara sampling.
- Cek apakah baris invalid ditangani dengan benar.
- Nilai insight berdasarkan data, bukan opini tanpa angka.

### Kebijakan Modul Berikutnya (MD Berikutnya):
- Jika ada function yang tidak tersedia di Excel 2013/2016, selalu beri notifikasi kompatibilitas.
- Jika function tersedia di Excel 2013 sampai versi terbaru, notifikasi kompatibilitas tidak perlu ditampilkan.
- Function yang tidak tersedia di Excel 2013/2016 tidak dimasukkan ke checklist penilaian utama.
- Function tersebut tetap boleh diberikan sebagai latihan mandiri di luar kelas.

---

## 📋 TEMPLATE CHECKLIST UNTUK COPY-PASTE

Mahasiswa dapat menggunakan template ini untuk track progress:

```
WEEK 7 - LOOKUP FUNCTIONS CHECKLIST

LEVEL 1 (Target: selesai 40 menit):
□ 1.1 - VLOOKUP Lengkapi Transaksi (15 min)
  ├─ Nama Produk via VLOOKUP
  ├─ Harga Unit via VLOOKUP
  ├─ Kategori via VLOOKUP
  └─ Invalid code ditangani IFERROR

□ 1.2 - Total Sales & Validasi (12 min)
  ├─ Total Sales = Qty x Harga
  ├─ Status Validasi (OK/Perlu Koreksi)
  └─ Semua baris terisi benar

□ 1.3 - HLOOKUP Target Bulanan (13 min)
  ├─ Lookup target sales bulanan
  ├─ Lookup target margin bulanan
  └─ Analisis gap target

LEVEL 2 (Target: selesai 50 menit):
□ 2.1 - Dashboard Kualitas Data (15 min)
  ├─ Total/Valid/Invalid transaksi
  ├─ Persentase valid data
  └─ KPI status data quality

□ 2.2 - Dashboard Sales Kategori (15 min)
  ├─ Ringkasan sales per kategori
  ├─ Identifikasi kategori dominan
  └─ Prioritas kategori promosi

□ 2.3 - Insight & Rekomendasi (20 min)
  ├─ Minimal 3 insight berbasis data
  ├─ Minimal 3 rekomendasi operasional
  └─ Setiap poin didukung angka

LATIHAN MANDIRI (OPSIONAL - TIDAK DINILAI):
□ INDEX + MATCH
  ├─ Lookup nama produk
  ├─ Lookup harga produk
  └─ Refleksi perbandingan dengan VLOOKUP

FINALISASI:
□ File naming: NAMA_NIM_Week7_LookupFunctions.xlsx
□ Semua formula visible
□ Struktur sheet rapi
□ Tidak ada error (#NAME?, #VALUE?, #N/A, #REF!)

Total Time: 90 menit
Passing Score: >= 70 (C3 >= 35 + C4 >= 35)
```

---

## End of Document
