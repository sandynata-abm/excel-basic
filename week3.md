# Latihan Pertemuan 3: Mathematical & Text Functions Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-003
- **Pertemuan**: 3
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Pemula (Beginner)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 (Basic) dan Week 2 (Formula)

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi matematika dan teks dalam konteks bisnis nyata
   - Menggunakan SUM, AVERAGE, MIN, MAX, COUNT untuk analisis data numerik
   - Mengimplementasikan LEFT, RIGHT, MID untuk manipulasi teks
   - Menggunakan CONCATENATE untuk menggabungkan field teks/data
   - Mengaplikasikan kombinasi fungsi untuk membuat laporan otomatis

2. **C4 (Analyze)** - Menganalisis data dengan perhitungan dan manipulasi teks
   - Membuat summary statistics (total, rata-rata, minimum, maksimum)
   - Mengekstrak informasi spesifik dari text fields
   - Membuat dashboard dengan perhitungan dan pemformatan teks otomatis
   - Menganalisis pola data untuk keputusan bisnis

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: Fungsi Agregasi Numerik (SUM, AVERAGE, MIN, MAX)

#### 1.1 SUM()
**Definisi**: Menjumlahkan semua angka dalam range tertentu.

**Sintaks**:
```
=SUM(range)
```

**Parameter**:
- range: Range sel yang berisi angka (contoh: A1:A10, atau A1:A5, B1:B5)

**Return Value**: Total penjumlahan semua nilai

**Kegunaan Bisnis**:
- Menghitung total penjualan/revenue
- Total biaya operasional
- Total unit inventory
- Total nilai invoice dalam periode tertentu

**Contoh**:
```
PT Maju Jaya - Total Penjualan Maret 2026
- Penjualan Minggu 1: Rp 1,500,000
- Penjualan Minggu 2: Rp 2,000,000
- Penjualan Minggu 3: Rp 1,750,000
- Penjualan Minggu 4: Rp 2,250,000

Formula: =SUM(A1:A4)
Hasil: Rp 7,500,000 (total penjualan bulanan)
```

**Karakteristik**:
- Mengabaikan sel kosong
- Hanya menjumlahkan angka (text diabaikan)
- Sangat efisien untuk range besar

---

#### 1.2 AVERAGE() (sering disebut AVG)
**Definisi**: Menghitung rata-rata (mean) dari angka dalam range.

**Sintaks**:
```
=AVERAGE(range)
```

**Parameter**:
- range: Range sel yang berisi angka

**Return Value**: Nilai rata-rata dari range

**Kegunaan Bisnis**:
- Menghitung rata-rata penjualan per hari/minggu/bulan
- Rata-rata harga jual
- Rata-rata jumlah unit terjual
- KPI monitoring (average order value, average transaction)

**Contoh**:
```
PT Maju Jaya - Rata-rata Penjualan Harian
- Penjualan 7 hari: 1500K, 2000K, 1750K, 2250K, 1800K, 2100K, 1900K

Formula: =AVERAGE(A1:A7)
Hasil: Rp 1,900,000 (rata-rata penjualan per hari)

Analisis: Hari-hari yang di bawah Rp 1,900K perlu ditingkatkan performanya
```

---

#### 1.3 MIN()
**Definisi**: Mencari nilai terkecil dalam range.

**Sintaks**:
```
=MIN(range)
```

**Parameter**:
- range: Range sel yang berisi angka

**Return Value**: Nilai terendah dari range

**Kegunaan Bisnis**:
- Mencari harga terendah (untuk negosiasi dengan supplier)
- Nilai transaksi terendah
- Inventory minimum alert
- Lowest performing product/salesman

**Contoh**:
```
PT Maju Jaya - Analisis Penjualan
Penjualan Harian: 1500K, 2000K, 1750K, 2250K, 1200K, 2100K, 1900K

Formula: =MIN(A1:A7)
Hasil: Rp 1,200,000 (penjualan terendah pada hari ke-5)

Interpretasi: Hari ke-5 underperform, perlu investigate alasannya
```

---

#### 1.4 MAX()
**Definisi**: Mencari nilai terbesar dalam range.

**Sintaks**:
```
=MAX(range)
```

**Parameter**:
- range: Range sel yang berisi angka

**Return Value**: Nilai tertinggi dari range

**Kegunaan Bisnis**:
- Mencari harga tertinggi
- Nilai transaksi terbesar
- Best performing salesman
- Peak sales day/period

**Contoh**:
```
PT Maju Jaya - Best Performance Monitoring
Penjualan Harian: 1500K, 2000K, 1750K, 2250K, 1200K, 2100K, 1900K

Formula: =MAX(A1:A7)
Hasil: Rp 2,250,000 (penjualan tertinggi pada hari ke-4)

Interpretasi: Hari ke-4 adalah peak day, replicate strategi hari tersebut
```

---

### Blok 2: Fungsi Penghitung (COUNT)

#### 2.1 COUNT()
**Definisi**: Menghitung jumlah sel yang berisi nilai numerik dalam range.

**Sintaks**:
```
=COUNT(range)
```

**Parameter**:
- range: Range sel yang akan dihitung

**Return Value**: Jumlah sel yang berisi angka

**Kegunaan Bisnis**:
- Menghitung jumlah transaksi yang berhasil
- Jumlah unit yang terjual
- Supply counter
- Missing value detection (jika COUNT < expected row count)

**Contoh**:
```
PT Maju Jaya - Transaction Counter
Data transaksi Maret 2026:
- Row 1 sampai 10: Data transaksi (berisi angka)
- Row 11: Sel kosong
- Row 12-15: Data transaksi

Formula: =COUNT(A1:A15)
Hasil: 14 (jumlah transaksi yang tercatat)
Catatan: Row 11 diabaikan karena kosong
```

**COUNT vs COUNTA**:
- COUNT: Hanya menghitung sel dengan angka
- COUNTA: Menghitung sel yang tidak kosong (termasuk text)
- Untuk praktikum ini, gunakan COUNT

---

### Blok 3: Fungsi Manipulasi Teks (LEFT, RIGHT, MID)

#### 3.1 LEFT()
**Definisi**: Mengekstrak karakter pertama (dari kiri) sejumlah N dari sebuah teks.

**Sintaks**:
```
=LEFT(text, num_chars)
```

**Parameter**:
- text: String teks yang akan diambil
- num_chars: Jumlah karakter yang ingin diambil dari kiri

**Return Value**: Karakter pertama sejumlah yang diminta

**Kegunaan Bisnis**:
- Ekstrak kode produk dari SKU (contoh: "PRDK-001" → "PRDK")
- Ambil inisial dari nama customer
- Extract area code dari nomor telepon
- Format ID untuk sorting/filtering

**Contoh**:
```
Skenario: PT Terpadu Trading - Ekstrak Kode Kategori Produk
Data: 
- SKU: "PRDK-001-RED"
- Kebutuhan: Ambil kategori produk (3 karakter pertama)

Formula: =LEFT(A1, 4)
Hasil: "PRDK" (kategori produk)

Contoh Lain:
- "STAT-2023-001" + LEFT(A2,4) = "STAT"
- "CUST-12345" + LEFT(A3,4) = "CUST"
```

---

#### 3.2 RIGHT()
**Definisi**: Mengekstrak karakter terakhir (dari kanan) sejumlah N dari sebuah teks.

**Sintaks**:
```
=RIGHT(text, num_chars)
```

**Parameter**:
- text: String teks yang akan diambil
- num_chars: Jumlah karakter yang ingin diambil dari kanan

**Return Value**: Karakter terakhir sejumlah yang diminta

**Kegunaan Bisnis**:
- Ekstrak nomor urut dari ID (contoh: "PRDK-001" → "001")
- Ambil extension dari email
- Extract area code dari format telepon
- Extract year/period dari date string

**Contoh**:
```
Skenario: PT Terpadu Trading - Ekstrak Invoice Number
Data:
- Invoice ID: "INV-2026-00456"
- Kebutuhan: Ambil nomor invoice (5 digit terakhir)

Formula: =RIGHT(A1, 5)
Hasil: "00456" (nomor urut invoice)

Contoh Lain:
- "PRDK-001" + RIGHT(A2,3) = "001"
- "invoice@company.com" + RIGHT(A3,3) = "com"
```

---

#### 3.3 MID()
**Definisi**: Mengekstrak sejumlah karakter dari tengah teks, mulai dari posisi tertentu.

**Sintaks**:
```
=MID(text, start_num, num_chars)
```

**Parameter**:
- text: String teks yang akan diambil
- start_num: Posisi karakter awal (mulai dari 1, bukan 0)
- num_chars: Jumlah karakter yang ingin diambil

**Return Value**: Karakter dari posisi middle sesuai jumlah yang diminta

**Kegunaan Bisnis**:
- Ekstrak bulan dari date string (format: "2026-03-29" → "03")
- Ekstrak departemen dari employee ID
- Extract middle part dari format code dengan pattern fixed

**Contoh**:
```
Skenario: PT Terpadu Trading - Ekstrak Bulan dari Tanggal
Data:
- Tanggal: "2026-03-29" (format: YYYY-MM-DD)
- Kebutuhan: Ambil bulan (posisi 6-7, 2 karakter)

Formula: =MID(A1, 6, 2)
Hasil: "03" (bulan Maret)

Penjelasan MID:
- Text: "2026-03-29"
- Posisi: 1=2, 2=0, 3=2, 4=6, 5=-, 6=0, 7=3, ...
- Start_num = 6 (dimulai dari "0")
- Num_chars = 2 (ambil "0" dan "3")
- Output: "03"

Contoh Lain:
- "EMP-SALES-001" + MID(A2, 5, 5) = "SALES"
- "2026-12-25" + MID(A3, 6, 2) = "12"
```

---

### Blok 4: Fungsi Penggabung Teks (CONCATENATE)

#### 4.1 CONCATENATE()
**Definisi**: Menggabungkan beberapa teks/nilai menjadi satu string.

**Sintaks - Metode 1 (Klasik)**:
```
=CONCATENATE(text1, text2, text3, ...)
```

**Sintaks - Metode 2 (Modern, dengan &)**:
```
=text1 & text2 & text3 & ...
```

**Parameter**:
- text1, text2, text3: Teks atau cell reference yang akan digabungkan

**Return Value**: String hasil penggabungan

**Kegunaan Bisnis**:
- Membuat full name dari first name + middle name + last name
- Membuat alamat lengkap dari street + city + province
- Membuat invoice ID otomatis (contoh: "INV-" & MONTH & YEAR & number)
- Membuat label untuk reporting

**Contoh**:
```
Skenario: PT Terpadu Trading - Buat Full Address
Data:
- Jalan: Jl. Merdeka
- Kota: Jakarta
- Provinsi: DKI Jakarta

Metode 1:
Formula: =CONCATENATE(A1, ", ", A2, ", ", A3)
Hasil: "Jl. Merdeka, Jakarta, DKI Jakarta"

Metode 2 (Lebih Sederhana):
Formula: =A1 & ", " & A2 & ", " & A3
Hasil: "Jl. Merdeka, Jakarta, DKI Jakarta"

Contoh Lain - Full Name:
- First Name: "Budi"
- Middle Name: "Santoso"
- Last Name: "Wijaya"
- Formula: =A1 & " " & A2 & " " & A3
- Hasil: "Budi Santoso Wijaya"
```

**Praktik Penggabungan dengan Format**:
```
Skenario: Buat Invoice ID Format "INV-[Bulan][Tahun]-[Nomor]"
- Bulan: 03 (March)
- Tahun: 26 (2026)
- Nomor: 001

Formula: ="INV-" & A1 & A2 & "-" & A3
Hasil: "INV-0326-001"

Atau dengan cell reference:
Formula: ="INV-" & B1 & B2 & "-" & REPT("0",3-LEN(B3)) & B3
Catatan: Ini advanced (lihat Week 4 untuk REPT, LEN)
```

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: Kasus Bisnis CV Komputer Indonesia

**Konteks Bisnis**:
CV Komputer Indonesia adalah perusahaan yang menjual perangkat komputer dan aksesori. Perusahaan memiliki 20 karyawan penjualan yang mencatat transaksi harian. Data ini merupakan catatan penjualan Maret 2026 yang perlu dianalisis untuk performance evaluation, reporting, dan inventory management.

### Struktur Kolom & Penjelasan

| No | Kolom | Tipe Data | Keterangan | Contoh |
|----|-------|-----------|-----------|---------|
| A | No | Integer | Nomor urut | 1, 2, 3, ... |
| B | Invoice ID | Text | Nomor invoice unique | INV-0326-001 |
| C | Nama Salesman | Text | Nama penjual | Budi Santoso |
| D | Produk | Text | Nama produk yang dijual | Monitor LG 24 inch |
| E | Kategori | Text | Kategori produk (extracted) | Electronics |
| F | Quantity | Integer | Jumlah unit terjual | 5 |
| G | Harga Per Unit | Currency | Harga satuan (Rp) | 1500000 |
| H | Total Penjualan | Currency | Subtotal = F×G | 7500000 |

**Data untuk Dianalisis**:
- Kolom A-D: Data input (dari sistem POS)
- Kolom E: Hasil extract dari D (menggunakan LEFT/MID)
- Kolom F-H: Data keuangan

### Sample Data Praktik (8 Transaksi)

```
No | Invoice ID  | Salesman       | Produk              | Qty | Harga Unit | Total
1  | INV-0326-001| Budi Santoso   | PRDK-Monitor-LG-001 | 2   | 1500000    | 3000000
2  | INV-0326-002| Ani Wijaya     | PRDK-Keyboard-001   | 5   | 500000     | 2500000
3  | INV-0326-003| Rudi Firmansah | PRDK-Mouse-001      | 10  | 150000     | 1500000
4  | INV-0326-004| Budi Santoso   | PRDK-Monitor-HD-002 | 1   | 2000000    | 2000000
5  | INV-0326-005| Ani Wijaya     | PRDK-RAM-8GB-001    | 3   | 600000     | 1800000
6  | INV-0326-006| Rudi Firmansah | PRDK-SSD-256-001    | 4   | 800000     | 3200000
7  | INV-0326-007| Budi Santoso   | PRDK-Monitor-LG-001 | 1   | 1500000    | 1500000
8  | INV-0326-008| Ani Wijaya     | PRDK-Keyboard-001   | 2   | 500000     | 1000000
```

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Functions
*Fokus: SUM, AVERAGE, MIN, MAX, COUNT - Durasi: 40 menit*

#### Latihan 1.1: Analisis Penjualan dengan SUM dan AVERAGE
**Waktu**: 12 menit

**Instruksi**:
1. Copy sample data di atas ke Excel (Kolom A sampai H, data 1-8)

2. Di bawah data, buat summary box:
   ```
   Total Penjualan: =SUM(H2:H9)
   Rata-rata Per Transaksi: =AVERAGE(H2:H9)
   Jumlah Transaksi: =COUNT(F2:F9)
   ```

3. **Analisis**:
   - Berapa total revenue Maret sampai saat ini?
   - Berapa rata-rata nilai transaksi?
   - Berapa banyak transaksi yang tercatat?

**Pertanyaan**:
- Jika target monthly revenue Rp 20 juta, sudah tercapai berapa persen?

---

#### Latihan 1.2: MIN dan MAX untuk Performance Analysis
**Waktu**: 8 menit

**Instruksi**:
1. Simulasikan data penjualan per salesman untuk 5 hari:
   ```
   Hari      | Budi   | Ani   | Rudi
   Hari 1    | 500K   | 750K  | 600K
   Hari 2    | 600K   | 800K  | 650K
   Hari 3    | 550K   | 700K  | 700K
   Hari 4    | 650K   | 850K  | 750K
   Hari 5    | 700K   | 900K  | 800K
   ```

2. Hitung:
   ```
   Penjualan Tertinggi: =MAX(range salesman)
   Penjualan Terendah: =MIN(range salesman)
   Penjualan Rata-rata: =AVERAGE(range salesman)
   ```

3. **Analisis**:
   - Siapa top performer (highest average)?
   - Siapa yang perlu improvement (lowest day)?

---

#### Latihan 1.3: COUNT untuk Data Validation
**Waktu**: 5 menit

**Instruksi**:
1. Gunakan data dari Latihan 1.1

2. Hitung:
   ```
   Total Record: =COUNT(A2:A9) atau =COUNTA(B2:B9)
   Transaksi dengan Qty > 0: =COUNT(F2:F9)
   ```

3. **Interpretasi**:
   - Apakah semua baris data lengkap?
   - Apakah ada data yang hilang atau error?

---

### LEVEL 2: Analisis Lanjutan dengan Text & Combined Functions
*Fokus: LEFT, RIGHT, MID, CONCATENATE - Durasi: 50 menit*

#### Latihan 2.1: Text Extraction dan Report Building
**Waktu**: 30 menit

**TAHAP A: Extract Kategori Produk (8 menit)**

1. Gunakan data dari Latihan 1.1 (Kolom D: Produk)
2. Contoh data: "PRDK-Monitor-LG-001"
3. Di kolom E (Kategori), buat formula untuk extract kategori:
   ```
   =LEFT(D2, 4)  → Hasil: "PRDK"
   ```
   Atau lebih detail (extract monitor type):
   ```
   =MID(D2, 6, 7)  → Hasil: "Monitor"
   ```

4. Drag formula ke semua row data

5. **Verifikasi**:
   - Apakah semua kategori ter-extract dengan benar?
   - Jika ada yang salah, adjust formula

---

**TAHAP B: Extract Invoice Number (7 menit)**

1. Gunakan data Kolom B (Invoice ID): "INV-0326-001"
2. Buat kolom baru "Invoice Number"
3. Formula untuk extract nomor invoice:
   ```
   =RIGHT(B2, 3)  → Hasil: "001"
   ```

4. Drag formula ke semua row

5. **Analisis**:
   - Nomor invoice yang paling tinggi? `=MAX(kolom_invoice_number)`
   - Ini membantu track sequential numbering

---

**TAHAP C: Create Full Report Label (15 menit)**

1. Buat kolom baru "Report Label" yang format:
   ```
   "Invoice INV-0326-001 | Budi Santoso menjual Monitor | Total: Rp 3.000.000"
   ```

2. Gunakan CONCATENATE atau & untuk menggabung:
   ```
   =C2 & " menjual " & D2 & " | Invoice: " & B2 & " | Total: Rp " & TEXT(H2,"0")
   ```
   Catatan: TEXT function untuk format currency (atau Week 4 topic)

3. Alternatif (lebih simple):
   ```
   ="Transaksi " & B2 & " | " & C2 & " | Qty: " & F2 & " unit"
   ```

4. Drag formula ke semua row

5. **Output**:
   - Setiap baris punya report label yang lengkap dan readable
   - Ini bisa digunakan untuk automated email/SMS notification, laporan harian, etc

---

#### Latihan 2.2: Dashboard dengan Summary & Extracted Data
**Waktu**: 20 menit

**Instruksi**:

1. **Buat Dashboard Area** (di samping data utama):

   ```
   DASHBOARD PENJUALAN - MARET 2026
   ================================
   
   A. SUMMARY METRICS:
   - Total Penjualan (SUM):        [formula]
   - Rata-rata Per Transaksi:      [formula]
   - Transaksi Tertinggi:          [formula]
   - Transaksi Terendah:           [formula]
   - Total Transaksi:              [formula]
   
   B. TOP PERFORMERS:
   - Nama Salesman dengan Revenue Tertinggi:  [formula]
   - Kategori Produk Terlaris:                [formula]
   ```

2. **Gunakan kombinasi formulas**:
   - SUM() untuk total
   - AVERAGE(), MAX(), MIN() untuk metrics
   - COUNT() untuk jumlah transaksi
   - INDEX/MATCH untuk mencari nama (atau simple VLOOKUP Week 4)

3. **Interpretasi Dashboard**:
   - Siapa best salesman?
   - Produk apa yang paling laris?
   - Apakah penjualan merata atau ada spike?

---

## 📋 INSTRUKSI PENGUMPULAN & EVALUASI

### Checklist Pengerjaan Praktikum (Durasi: 90 Menit)

**LEVEL 1 - Penerapan Dasar (40 menit)**:
- [ ] Latihan 1.1: SUM, AVERAGE, COUNT untuk summary (12 menit)
- [ ] Latihan 1.2: MIN, MAX untuk performance analysis (8 menit)
- [ ] Latihan 1.3: COUNT untuk data validation (5 menit)
- [ ] Buffer/Review: 15 menit

**LEVEL 2 - Analisis Lanjutan (50 menit)**:
- [ ] Latihan 2.1 TAHAP A: LEFT/MID untuk ekstrak kategori (8 menit)
- [ ] Latihan 2.1 TAHAP B: RIGHT untuk ekstrak invoice number (7 menit)
- [ ] Latihan 2.1 TAHAP C: CONCATENATE untuk report label (15 menit)
- [ ] Latihan 2.2: Dashboard dengan summary metrics (20 menit)

**Finalisasi (10 menit min)**:
- [ ] File saved dengan naming: `NAMA_NIM_Week3_MathTextFunction.xlsx`
- [ ] Setiap worksheet diberi title dan keterangan kolom
- [ ] Formulas ter-dokumentasi (keterangan di samping/bawah)

---

### Rubric Penilaian (Outcome Based Education)

| Learning Outcome | Indikator Ketercapaian | Skor |
|-----------------|-------------------------|------|
| **C3 (Apply)** - Menerapkan Math & Text functions | Semua formula SUM, AVERAGE, MIN, MAX, COUNT, LEFT, RIGHT, MID, CONCATENATE diinput dengan benar; output sesuai harapan; tanpa error | 50% |
| **C4 (Analyze)** - Menganalisis data dengan formulas | Dashboard dibuat dengan summary metrics yang tepat; ekstraksi text benar; interpretasi business insight jelas; report label informatif | 50% |

**Kriteria Passing**: Score ≥ 70 (C3 ≥ 35 + C4 ≥ 35)



---

## 📚 TIPS & TROUBLESHOOTING PRAKTIKUM

### Common Mistakes & Fixes

| Masalah | Solusi |
|---------|--------|
| SUM menghasilkan 0 padahal ada data | Cek: apakah range benar? A1:A10 bukan A1:B10 (double column). Data numeric bukan text. Format sel adalah Number bukan Text. |
| AVERAGE menghasilkan #DIV/0! | Range kosong atau semua text. Cek: apakah ada angka dalam range? Jika all text, gunakan COUNT(range) dulu. |
| LEFT/RIGHT tidak berfungsi | Cek parameter: LEFT(text, num_chars) - pastikan text dari cell reference atau quoted string, num_chars adalah angka. |
| MID menghasilkan hasil kosong | Start_num atau num_chars mungkin lebih besar dari panjang text. Test dengan text lebih panjang atau adjust start_num. |
| CONCATENATE error #VALUE! | Pastikan menggunakan & antar text, bukan +. Contoh BENAR: =A1 & B1, SALAH: =A1 + B1. |
| Formula tidak bisa di-drag dengan benar | Cell reference harus relative (A1) bukan absolute ($A$1). Saat drag, A1 akan jadi A2, A3, dst otomatis. |

### Shortcut & Tips Efisiensi

1. **SUM Range Quick**: Click sel awal, hold Shift, click sel akhir. Langsung double-click fill handle untuk extend.
2. **Format Numeric**: Jika SUM = 0 padahal data ada, pilih range → Format Cells → Number (bukan Text).
3. **Text Function Testing**: Test di cell terpisah dulu sebelum drag ke banyak row.
4. **CONCATENATE Alternative**: Gunakan & operator (lebih cepat dari CONCATENATE function).
   - Benar: =A1 & "-" & B1
   - Salah: =CONCATENATE(A1, "-", B1) ← lebih panjang
5. **Extract Formula Building**: Jika tidak yakin posisi, gunakan Find & Replace untuk test pattern terlebih dahulu.

---

## 🔧 INSTRUKSI UNTUK AGENT: HTML GENERATION

### Purpose
Markdown ini adalah blueprint untuk generate HTML file yang professional, accessible, dan sesuai Tailwind CSS styling.

### Key Points untuk Agent

1. **Layout & Structure**
   - Gunakan 3-column grid (sidebar + main content) seperti week1.html & week2.html
   - Sidebar (1/3): Materi list + Tips
   - Main content (2/3): Panduan + Tabel data + Instruksi tugas

2. **Color Scheme**
   - Header background: #217346 (Excel green)
   - Accent: green-600, green-500, blue-500 (untuk code blocks)
   - Neutral: gray-50, gray-100, gray-200, gray-800

3. **Component Reusability**
   - Header component (icon + title + subtitle)
   - Material list (use Font Awesome icons)
   - Tips box (green background)
   - Definition box (gray background dengan left border)
   - Exercise card (dengan checkbox untuk progress tracking)
   - Table (standard Tailwind table dengan border)
   - Footer component (copyright + identitas institusi, konsisten dengan week1.html/week2.html)
   - Code block dengan syntax highlighting (pre/code tags)

4. **Interactive Features**
   - Copy table to clipboard button (untuk praktikum)
   - Pastikan data nominal/bilangan yang di-copy menggunakan format Number (bukan Text) agar saat paste ke Excel tetap terdeteksi sebagai angka
   - Collapsible exercises (optional, untuk UX yang lebih baik)
   - Progress tracker (visual checklist)
   - Formula copy button untuk memudahkan student

5. **Responsiveness**
   - Desktop: 3-column layout
   - Tablet: 2-column (sidebar moves down)
   - Mobile: Full-width single column

---

## ✅ TEMPLATE CHECKLIST UNTUK EXERCISE BERIKUTNYA

Gunakan struktur ini ketika membuat latihan minggu berikutnya:

- [ ] **Metadata & OBE Framework**: Learning outcomes yang jelas (C3, C4, min 4 outcomes)
- [ ] **Materi Terstruktur**: Min 3-4 blok, setiap blok minimal 2 fungsi/konsep
- [ ] **Business Context**: Praktik relevan dengan dunia bisnis/ekonomi
- [ ] **Sample Data**: Realistic business data, 5-10 rows (atau 8 rows)
- [ ] **Latihan Bertingkat**: LEVEL 1 (40 min) dan LEVEL 2 (50 min)
- [ ] **Dashboard/Analysis**: Minimal 1 dashboard dengan summary metrics
- [ ] **Extraction/Manipulation**: Text processing atau data manipulation lessons
- [ ] **Documentation**: Jelas, professional, mudah dipahami
- [ ] **Rubric Penilaian**: OBE-aligned dengan C3 (Apply) + C4 (Analyze)

---

**End of Document**

*Dokumen ini adalah template comprehensive untuk latihan Excel Mathematical & Text Functions. 
Dipersiapkan untuk Outcome Based Education, business-context relevant, 
dan AI-friendly untuk code generation.*
