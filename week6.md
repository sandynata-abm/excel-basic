# Latihan Pertemuan 6: Criteria-Based Functions Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-006
- **Pertemuan**: 6
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Pemula (Beginner)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 s.d. Week 5

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi berbasis kriteria dalam konteks bisnis nyata
   - Menggunakan SUMIF() untuk penjumlahan berdasarkan satu kriteria
   - Menggunakan COUNTIF() untuk menghitung jumlah data berdasarkan satu kriteria
   - Menggunakan AVERAGEIF() untuk menghitung rata-rata berdasarkan satu kriteria
   - Menggabungkan fungsi kriteria dengan dashboard metrik operasional

2. **C4 (Analyze)** - Menganalisis data berdasarkan segmentasi kriteria
   - Menganalisis performa penjualan per kategori produk
   - Menganalisis distribusi transaksi per status/region/channel
   - Menyusun insight bisnis berbasis metrik segmentasi
   - Merumuskan rekomendasi operasional dari data terfilter

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: SUMIF() - Penjumlahan Berdasarkan Satu Kriteria

#### 1.1 SUMIF() Dasar
**Definisi**: Menjumlahkan nilai pada range tertentu yang memenuhi satu kriteria.

**Sintaks**:
```
=SUMIF(range_kriteria, kriteria, [range_penjumlahan])
```

**Parameter**:
- range_kriteria: Range yang diuji dengan kriteria
- kriteria: Syarat (angka, teks, operator, atau referensi cell)
- range_penjumlahan (opsional): Range nilai yang dijumlahkan

**Return Value**: Total nilai yang memenuhi kriteria.

**Kegunaan Bisnis**:
- Total penjualan untuk satu kategori produk
- Total transaksi untuk satu region
- Total revenue untuk channel tertentu

**Contoh**:
```
CV Komputer Indonesia - Sales by Category
- Tujuan: Total penjualan kategori "Laptop"
- Formula: =SUMIF(C2:C9, "Laptop", G2:G9)
- Hasil: Total revenue seluruh transaksi kategori Laptop

Interpretasi: Mengetahui kontribusi revenue per kategori produk
```

**Catatan Praktis**:
- Jika range_penjumlahan dikosongkan, Excel akan menjumlahkan range_kriteria.
- Gunakan wildcard seperti "*Laptop*" jika butuh pencocokan parsial teks.

---

#### 1.2 SUMIF() dengan Operator Logika
**Definisi**: SUMIF dengan syarat numerik menggunakan operator.

**Contoh Kriteria**:
- ">=50000000" (nilai lebih dari atau sama dengan 50 juta)
- "<1000000" (nilai kurang dari 1 juta)

**Contoh Formula**:
```
=SUMIF(G2:G9, ">=50000000", G2:G9)
```

**Interpretasi**:
Menjumlahkan seluruh transaksi bernilai besar untuk analisis high-value sales.

---

### Blok 2: COUNTIF() - Penghitungan Berdasarkan Satu Kriteria

#### 2.1 COUNTIF() Dasar
**Definisi**: Menghitung jumlah cell yang memenuhi satu kriteria.

**Sintaks**:
```
=COUNTIF(range, kriteria)
```

**Parameter**:
- range: Range yang diperiksa
- kriteria: Syarat perhitungan

**Return Value**: Jumlah item yang sesuai kriteria.

**Kegunaan Bisnis**:
- Menghitung jumlah transaksi dengan status tertentu
- Menghitung jumlah transaksi per region
- Menghitung jumlah produk per kategori

**Contoh**:
```
CV Komputer Indonesia - Order Status Count
- Tujuan: Hitung jumlah transaksi status "Lunas"
- Formula: =COUNTIF(H2:H9, "Lunas")
- Hasil: Jumlah transaksi lunas

Interpretasi: Memantau kualitas cashflow berdasarkan status pembayaran
```

---

#### 2.2 COUNTIF() dengan Operator dan Wildcard
**Contoh Formula**:
```
=COUNTIF(G2:G9, ">=30000000")
=COUNTIF(D2:D9, "*Online*")
```

**Interpretasi**:
- Formula pertama menghitung transaksi dengan nilai minimum 30 juta.
- Formula kedua menghitung transaksi dengan channel yang mengandung teks "Online".

---

### Blok 3: AVERAGEIF() - Rata-Rata Berdasarkan Satu Kriteria

#### 3.1 AVERAGEIF() Dasar
**Definisi**: Menghitung nilai rata-rata dari data yang memenuhi satu kriteria.

**Sintaks**:
```
=AVERAGEIF(range_kriteria, kriteria, [range_rata_rata])
```

**Parameter**:
- range_kriteria: Range yang diuji kriterianya
- kriteria: Syarat filter
- range_rata_rata (opsional): Range nilai yang dirata-rata

**Return Value**: Nilai rata-rata sesuai kriteria.

**Kegunaan Bisnis**:
- Rata-rata penjualan per kategori
- Rata-rata nilai order per channel
- Rata-rata transaksi untuk region tertentu

**Contoh**:
```
CV Komputer Indonesia - Average Sales by Region
- Tujuan: Rata-rata nilai transaksi region "Malang"
- Formula: =AVERAGEIF(B2:B9, "Malang", G2:G9)
- Hasil: Average revenue transaksi region Malang

Interpretasi: Menilai kualitas rata-rata transaksi per wilayah
```

---

#### 3.2 AVERAGEIF() untuk Kategori Operasional
**Contoh Formula**:
```
=AVERAGEIF(C2:C9, "Printer", G2:G9)
```

**Interpretasi**:
Mengetahui rata-rata nominal penjualan untuk kategori Printer.

---

### Blok 4: Integrasi SUMIF + COUNTIF + AVERAGEIF untuk Dashboard

#### 4.1 Dashboard Ringkas Segmentasi
**Definisi**: Menggabungkan tiga fungsi utama untuk menghasilkan insight cepat.

**Contoh Metrik Dashboard**:
- Total Sales Kategori Laptop: `=SUMIF(C2:C9, "Laptop", G2:G9)`
- Jumlah Order Lunas: `=COUNTIF(H2:H9, "Lunas")`
- Rata-rata Sales Region Malang: `=AVERAGEIF(B2:B9, "Malang", G2:G9)`

**Kegunaan Bisnis**:
- Monitoring performa berdasarkan kriteria utama
- Membuat KPI operasional harian/mingguan
- Mendukung keputusan stok, promosi, dan collection

---

### Blok 5 (Opsional): SUMIFS, COUNTIFS, AVERAGEIFS - Belajar Mandiri

**Status Materi**:
- Materi ini **opsional** untuk eksplorasi mandiri di luar kelas/perkuliahan.
- Tidak masuk checklist penilaian utama pada praktikum kelas.

#### 5.1 SUMIFS()
**Definisi**: Menjumlahkan nilai berdasarkan banyak kriteria.

**Sintaks**:
```
=SUMIFS(range_penjumlahan, range_kriteria1, kriteria1, [range_kriteria2, kriteria2], ...)
```

**Contoh**:
```
=SUMIFS(G2:G9, C2:C9, "Laptop", H2:H9, "Lunas")
```
Interpretasi: Total sales Laptop yang statusnya Lunas.

#### 5.2 COUNTIFS()
**Definisi**: Menghitung jumlah data berdasarkan banyak kriteria.

**Sintaks**:
```
=COUNTIFS(range_kriteria1, kriteria1, [range_kriteria2, kriteria2], ...)
```

**Contoh**:
```
=COUNTIFS(B2:B9, "Malang", H2:H9, "Lunas")
```
Interpretasi: Jumlah order region Malang dengan status Lunas.

#### 5.3 AVERAGEIFS()
**Definisi**: Menghitung rata-rata nilai berdasarkan banyak kriteria.

**Sintaks**:
```
=AVERAGEIFS(range_rata_rata, range_kriteria1, kriteria1, [range_kriteria2, kriteria2], ...)
```

**Contoh**:
```
=AVERAGEIFS(G2:G9, C2:C9, "Monitor", H2:H9, "Lunas")
```
Interpretasi: Rata-rata transaksi Monitor yang sudah Lunas.

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: CV Komputer Indonesia - Sales Transaction Segment

**Konteks Bisnis**:
Tim operasional CV Komputer Indonesia ingin menganalisis transaksi berdasarkan kategori produk, region, channel, dan status pembayaran untuk menentukan prioritas promosi, stok, dan collection.

### Struktur Kolom & Penjelasan

| No | Kolom | Tipe Data | Keterangan | Contoh |
|----|-------|-----------|-----------|---------|
| A | No | Integer | Nomor urut transaksi | 1 |
| B | Region | Text | Wilayah transaksi | Malang |
| C | Kategori | Text | Kategori produk | Laptop |
| D | Channel | Text | Sumber order | Online |
| E | Qty | Integer | Jumlah unit | 3 |
| F | Harga Unit | Currency | Harga per unit (Rp) | 8,500,000 |
| G | Total Sales | Currency | Nilai total transaksi (Rp) | 25,500,000 |
| H | Payment Status | Text | Status pembayaran | Lunas |
| I | Priority Flag | Text | Flag operasional (calculated) | High |

**Data untuk Dianalisis**:
- Kolom A-H: Data input
- Kolom I: Hasil rule IF (prioritas)

### Sample Data Praktik (8 Transaksi)

```
No | Region  | Kategori | Channel | Qty | Harga Unit | Total Sales | Payment Status | Priority Flag
1  | Malang  | Laptop   | Online  | 3   | 8500000    | 25500000    | Lunas          | [Formula]
2  | Surabaya| Printer  | Offline | 5   | 2200000    | 11000000    | Lunas          | [Formula]
3  | Malang  | Monitor  | Online  | 4   | 3200000    | 12800000    | Hutang         | [Formula]
4  | Blitar  | Laptop   | Offline | 2   | 9200000    | 18400000    | Lunas          | [Formula]
5  | Kediri  | Aksesoris| Online  | 20  | 350000     | 7000000     | Lunas          | [Formula]
6  | Malang  | Printer  | Online  | 6   | 2100000    | 12600000    | Hutang         | [Formula]
7  | Surabaya| Laptop   | Online  | 7   | 8700000    | 60900000    | Lunas          | [Formula]
8  | Malang  | Monitor  | Offline | 3   | 3400000    | 10200000    | Lunas          | [Formula]
```

**Keterangan Data**:
- Total Sales sudah dihitung dari Qty x Harga Unit.
- Priority Flag akan dipakai untuk menandai transaksi prioritas follow-up.

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Functions
*Fokus: SUMIF, COUNTIF, AVERAGEIF - Durasi: 40 menit*

#### Latihan 1.1: Total Sales per Kategori dengan SUMIF
**Waktu**: 14 menit

**Instruksi**:
1. Copy sample data ke Excel (Kolom A sampai H).
2. Buat ringkasan total sales per kategori:
   ```
   Laptop   : =SUMIF(C2:C9, "Laptop", G2:G9)
   Printer  : =SUMIF(C2:C9, "Printer", G2:G9)
   Monitor  : =SUMIF(C2:C9, "Monitor", G2:G9)
   ```
3. Tambahkan total sales transaksi bernilai >= 15,000,000:
   ```
   =SUMIF(G2:G9, ">=15000000", G2:G9)
   ```

**Analisis**:
- Kategori mana dengan kontribusi sales terbesar?
- Berapa porsi high-value transaction terhadap total sales?

---

#### Latihan 1.2: Hitung Distribusi Data dengan COUNTIF
**Waktu**: 12 menit

**Instruksi**:
1. Hitung jumlah transaksi berdasarkan payment status:
   ```
   Lunas : =COUNTIF(H2:H9, "Lunas")
   Hutang: =COUNTIF(H2:H9, "Hutang")
   ```
2. Hitung jumlah transaksi online:
   ```
   =COUNTIF(D2:D9, "Online")
   ```
3. Hitung jumlah transaksi dengan total sales >= 20,000,000:
   ```
   =COUNTIF(G2:G9, ">=20000000")
   ```

**Analisis**:
- Bagaimana komposisi transaksi lunas vs hutang?
- Apakah channel online mendominasi?

---

#### Latihan 1.3: Rata-Rata Segmentasi dengan AVERAGEIF
**Waktu**: 14 menit

**Instruksi**:
1. Hitung rata-rata total sales untuk region Malang:
   ```
   =AVERAGEIF(B2:B9, "Malang", G2:G9)
   ```
2. Hitung rata-rata total sales untuk kategori Printer:
   ```
   =AVERAGEIF(C2:C9, "Printer", G2:G9)
   ```
3. Hitung rata-rata total sales transaksi lunas:
   ```
   =AVERAGEIF(H2:H9, "Lunas", G2:G9)
   ```

**Analisis**:
- Region mana yang sebaiknya diprioritaskan untuk campaign?
- Kategori apa dengan nilai transaksi rata-rata paling tinggi?

---

### LEVEL 2: Analisis Lanjut & Dashboard Segmentasi
*Fokus: Integrasi IF + SUMIF/COUNTIF/AVERAGEIF - Durasi: 50 menit*

#### Latihan 2.1: Priority Flag Rule dengan IF
**Waktu**: 12 menit

**Instruksi**:
1. Tambah kolom I dengan logika prioritas:
   ```
   Rule:
   - High  : Total Sales >= 20000000 ATAU Payment Status = "Hutang"
   - Medium: Total Sales >= 12000000
   - Low   : Selain itu
   
   Formula:
   =IF(OR(G2>=20000000, H2="Hutang"), "High", IF(G2>=12000000, "Medium", "Low"))
   ```
2. Copy formula ke seluruh data.

**Analisis**:
- Berapa transaksi berstatus High?
- Apakah hutang otomatis masuk prioritas follow-up?

---

#### Latihan 2.2: KPI Summary Panel
**Waktu**: 18 menit

**Instruksi**:
1. Buat KPI panel berikut:
   ```
   A. SALES KPI
   - Total Sales Semua Data: =SUM(G2:G9)
   - Total Sales Laptop: =SUMIF(C2:C9, "Laptop", G2:G9)
   - Total Sales Online: =SUMIF(D2:D9, "Online", G2:G9)

   B. COUNT KPI
   - Jumlah Transaksi Lunas: =COUNTIF(H2:H9, "Lunas")
   - Jumlah Transaksi Hutang: =COUNTIF(H2:H9, "Hutang")
   - Jumlah Priority High: =COUNTIF(I2:I9, "High")

   C. AVERAGE KPI
   - Avg Sales Malang: =AVERAGEIF(B2:B9, "Malang", G2:G9)
   - Avg Sales Laptop: =AVERAGEIF(C2:C9, "Laptop", G2:G9)
   - Avg Sales Lunas: =AVERAGEIF(H2:H9, "Lunas", G2:G9)
   ```
2. Tambahkan kolom status KPI dengan IF:
   ```
   =IF([nilai KPI] >= [target], "ON TRACK", "BELOW TARGET")
   ```

**Analisis**:
- KPI mana yang sudah memenuhi target?
- Area mana yang butuh perbaikan cepat?

---

#### Latihan 2.3: Insight & Rekomendasi Operasional
**Waktu**: 20 menit

**Instruksi**:
1. Tulis minimal 3 insight berbasis hasil fungsi criteria-based.
2. Tulis minimal 3 rekomendasi operasional (stok, promosi, collection).
3. Sertakan pembuktian angka dari KPI panel.

**Contoh Format Insight**:
```
Insight 1: Kategori Laptop menyumbang revenue tertinggi sebesar [angka],
sehingga layak menjadi fokus promosi utama.

Insight 2: Transaksi Hutang berjumlah [angka], dengan [angka] termasuk Priority High,
perlu follow-up collection maksimal 3 hari kerja.
```

---

## 📘 LATIHAN MANDIRI (OPSIONAL - DI LUAR KELAS)

### Topik: SUMIFS, COUNTIFS, AVERAGEIFS

**Status**:
- Tidak masuk checklist penilaian utama kelas.
- Direkomendasikan untuk eksplorasi kemampuan analisis multi-kriteria.

### Tugas Mandiri 1: Sales by Multi-Criteria

Gunakan formula berikut:
```
Total Sales Laptop Lunas:
=SUMIFS(G2:G9, C2:C9, "Laptop", H2:H9, "Lunas")

Jumlah Transaksi Online Malang:
=COUNTIFS(D2:D9, "Online", B2:B9, "Malang")

Rata-rata Sales Monitor Lunas:
=AVERAGEIFS(G2:G9, C2:C9, "Monitor", H2:H9, "Lunas")
```

### Tugas Mandiri 2: Buat 2 Formula Baru

1. Buat 1 formula SUMIFS sesuai kebutuhan analisis Anda.
2. Buat 1 formula COUNTIFS atau AVERAGEIFS sesuai kebutuhan Anda.
3. Tulis interpretasi hasilnya dalam 2-3 kalimat.

---

## 📋 INSTRUKSI PENGUMPULAN

### Format File
- Nama file: **NAMA_NIM_Week6_CriteriaFunctions.xlsx**
- Sheet utama: Data + Formula
- Sheet dashboard: KPI Summary (recommended)
- Semua formula visible (tidak di-hide)

### Checklist Pengerjaan (90 menit):
- [ ] LEVEL 1 (40 min): Latihan 1.1, 1.2, 1.3 selesai tanpa error
- [ ] LEVEL 2 (50 min): Latihan 2.1, 2.2, 2.3 selesai + insight
- [ ] Dashboard KPI: Sales, Count, Average metrics lengkap
- [ ] File save sesuai format nama

### Catatan Checklist:
- Latihan SUMIFS, COUNTIFS, AVERAGEIFS bersifat mandiri (opsional) dan tidak dinilai dalam checklist utama kelas.

### Rubric Penilaian:

| Aspek | Kriteria | Bobot |
|-------|----------|-------|
| **C3 (Apply)** | Formula SUMIF, COUNTIF, AVERAGEIF, IF diinput benar; hasil perhitungan akurat | 50% |
| **C4 (Analyze)** | KPI dashboard dan insight operasional jelas, berbasis data, dan relevan | 50% |

**Passing**: Score >= 70 (C3 >= 35 + C4 >= 35)

---

## 💡 TIPS & TROUBLESHOOTING

### Common Mistakes:

1. **Range Ukuran Tidak Sama**:
   ```
   ❌ SALAH: =SUMIF(C2:C9, "Laptop", G2:G10)
   ✅ BENAR: =SUMIF(C2:C9, "Laptop", G2:G9)
   ```

2. **Kriteria Teks Tidak Konsisten**:
   ```
   ❌ SALAH: =COUNTIF(H2:H9, "lunas")  [jika data menggunakan "Lunas"]
   ✅ BENAR: =COUNTIF(H2:H9, "Lunas")
   ```

3. **Operator Kriteria Salah Penulisan**:
   ```
   ❌ SALAH: =COUNTIF(G2:G9, >=20000000)
   ✅ BENAR: =COUNTIF(G2:G9, ">=20000000")
   ```

4. **AVERAGEIF Menghasilkan #DIV/0!**:
   - Penyebab: Tidak ada data yang memenuhi kriteria.
   - Solusi: Pastikan kriteria ada di data atau gunakan IFERROR.

5. **SUMIF Tidak Sesuai Ekspektasi**:
   - Cek spasi tersembunyi pada data teks.
   - Gunakan TRIM pada data sumber jika perlu.

### Troubleshooting Tips:
- Jika hasil 0 padahal harusnya ada nilai: cek typo kriteria.
- Jika hasil terlalu besar/kecil: cek range formula tertukar.
- Jika dashboard sulit dibaca: kelompokkan KPI per tema (Sales, Count, Average).

### Function Alternatives:
- Single criteria -> gunakan SUMIF / COUNTIF / AVERAGEIF.
- Multi criteria -> gunakan SUMIFS / COUNTIFS / AVERAGEIFS (latihan mandiri).
- Penanganan error -> kombinasikan IFERROR.

---

## 📖 INSTRUKSI UNTUK AGENT / TEACHING ASSISTANT

### Learning Objectives (Detailed):

1. **Criteria Function Mastery**:
   - Mahasiswa paham beda SUMIF, COUNTIF, AVERAGEIF.
   - Mahasiswa mampu menerapkan kriteria teks dan numerik.
   - Mahasiswa mampu memvalidasi hasil formula dengan logika bisnis.

2. **Business Interpretation**:
   - Menerjemahkan output formula menjadi insight operasional.
   - Memprioritaskan tindak lanjut berdasarkan payment status dan nilai transaksi.

3. **Dashboard Thinking**:
   - Menyusun KPI panel yang ringkas dan action-oriented.
   - Menentukan target KPI dan status ON TRACK/BELOW TARGET.

### Common Misconceptions to Clarify:

| Misconception | Correction |
|---|---|
| SUMIF dan SUMIFS sama | SUMIF untuk 1 kriteria, SUMIFS untuk banyak kriteria |
| COUNTIF bisa menghitung total nilai | COUNTIF hanya menghitung jumlah item, bukan menjumlahkan nominal |
| AVERAGEIF selalu aman | Bisa #DIV/0! jika tidak ada data yang match |
| Kriteria angka tidak perlu tanda petik | Pada operator (>=, <=, <, >) harus dalam string |
| Materi opsional tidak penting | Opsional tetap penting untuk peningkatan kompetensi mandiri |

### Assessment Tips:
- Uji formula dengan 1-2 data manual untuk validasi.
- Pastikan range kriteria dan range hasil sejajar ukuran.
- Minta mahasiswa jelaskan arti bisnis dari setiap KPI, bukan hanya angka.

### Kebijakan Modul Berikutnya (MD Berikutnya):
- Jika ada function yang tidak tersedia di Excel 2013/2016, selalu beri notifikasi kompatibilitas.
- Jika function tersedia di Excel 2013 sampai versi terbaru, notifikasi kompatibilitas tidak perlu ditampilkan.
- Function yang tidak tersedia di Excel 2013/2016 tidak dimasukkan ke checklist penilaian utama.
- Function tersebut tetap boleh diberikan sebagai latihan mandiri di luar kelas.

---

## 📋 TEMPLATE CHECKLIST UNTUK COPY-PASTE

Mahasiswa dapat menggunakan template ini untuk track progress:

```
WEEK 6 - CRITERIA FUNCTIONS CHECKLIST

LEVEL 1 (Target: selesai 40 menit):
□ 1.1 - SUMIF per Kategori (14 min)
  ├─ Formula SUMIF untuk 3 kategori
  ├─ Formula SUMIF dengan operator >=
  └─ Analisis kontribusi sales

□ 1.2 - COUNTIF Distribusi Transaksi (12 min)
  ├─ Count status Lunas/Hutang
  ├─ Count channel Online
  └─ Count transaksi >= 20jt

□ 1.3 - AVERAGEIF Segmentasi (14 min)
  ├─ Avg per Region
  ├─ Avg per Kategori
  └─ Avg berdasarkan Payment Status

LEVEL 2 (Target: selesai 50 menit):
□ 2.1 - Priority Flag IF (12 min)
  ├─ Rule High/Medium/Low
  ├─ Formula IF + OR + nested IF
  └─ Validasi seluruh baris data

□ 2.2 - KPI Summary Panel (18 min)
  ├─ Sales KPI (3 metrik)
  ├─ Count KPI (3 metrik)
  ├─ Average KPI (3 metrik)
  └─ Status KPI (ON TRACK/BELOW TARGET)

□ 2.3 - Insight & Rekomendasi (20 min)
  ├─ Minimal 3 insight
  ├─ Minimal 3 rekomendasi
  └─ Berbasis angka KPI

LATIHAN MANDIRI (OPSIONAL - TIDAK DINILAI):
□ SUMIFS / COUNTIFS / AVERAGEIFS
  ├─ Eksplorasi multi-kriteria
  ├─ Buat minimal 2 formula tambahan
  └─ Tulis interpretasi singkat hasil

FINALISASI:
□ File naming: NAMA_NIM_Week6_CriteriaFunctions.xlsx
□ Semua formula visible
□ Struktur sheet rapi
□ Tidak ada error (#NAME?, #VALUE?, #DIV/0!)

Total Time: 90 menit
Passing Score: >= 70 (C3 >= 35 + C4 >= 35)
```

---

## End of Document
