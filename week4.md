# Latihan Pertemuan 4: Date-Time Functions Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-004
- **Pertemuan**: 4
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Pemula (Beginner)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 (Basic), Week 2 (Formula), dan Week 3 (Mathematical & Text Functions)

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi tanggal dan waktu dalam konteks bisnis nyata
   - Menggunakan TODAY(), NOW() untuk mencatat timestamp otomatis
   - Mengimplementasikan YEAR(), MONTH(), DAY() untuk ekstrak komponen tanggal
   - Menggunakan DATEDIF untuk menghitung selisih hari/bulan/tahun
   - Mengaplikasikan NETWORKDAYS(), WORKDAY() untuk perhitungan hari kerja
   - Menggunakan EDATE() untuk manipulasi tanggal

2. **C4 (Analyze)** - Menganalisis data temporal dan membuat timeline reporting
   - Membuat employee aging report berdasarkan hire date
   - Menghitung masa kerja dan tenure analysis
   - Merencanakan deadline dan project timeline dengan WORKDAY
   - Menganalisis lead time dan project duration
   - Membuat HR analytics dan workforce planning

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: Fungsi Tanggal Dasar (TODAY, NOW, YEAR, MONTH, DAY)

#### 1.1 TODAY()
**Definisi**: Mengembalikan tanggal hari ini (sistem komputer) tanpa komponen waktu.

**Sintaks**:
```
=TODAY()
```

**Parameter**:
- Tidak ada parameter

**Return Value**: Tanggal dalam format date (misal: 2026-03-29)

**Kegunaan Bisnis**:
- Mencatat tanggal transaksi otomatis
- Menghitung umur dokumen (document age)
- Membuat timestamp untuk audit trail
- Menghitung age of receivables (usia piutang)

**Contoh**:
```
CV Komputer Indonesia - Aging Report
- Invoice dibuat tanggal 2026-03-01
- Hari ini: =TODAY() → 2026-03-29
- Usia invoice: =TODAY() - D2 → 28 hari

Interpretasi: Invoice berusia 28 hari, harus segera ditagih
```

**Karakteristik**:
- Otomatis update setiap hari
- Tidak perlu input manual
- Tidak termasuk jam/menit/detik

---

#### 1.2 NOW()
**Definisi**: Mengembalikan tanggal dan waktu saat ini (sistem komputer).

**Sintaks**:
```
=NOW()
```

**Parameter**:
- Tidak ada parameter

**Return Value**: Tanggal + waktu (misal: 2026-03-29 14:35:22)

**Kegunaan Bisnis**:
- Mencatat timestamp lengkap untuk log aktivitas
- Tracking waktu entry/exit, clock-in/clock-out
- Menghitung elapsed time (waktu berlalu)
- Audit timestamp untuk compliance

**Contoh**:
```
CV Komputer Indonesia - Transaction Log
- Transaksi masuk: =NOW() → 2026-03-29 14:35:22
- Processing time: =NOW() - log_time

Interpretasi: Data transaksi tersimpan dengan timestamp presisi
```

---

#### 1.3 YEAR()
**Definisi**: Mengekstrak tahun dari sebuah tanggal.

**Sintaks**:
```
=YEAR(date)
```

**Parameter**:
- date: Nilai tanggal (bisa cell reference atau TODAY(), NOW())

**Return Value**: Angka tahun (misal: 2026)

**Kegunaan Bisnis**:
- Mengelompokkan data berdasarkan tahun (year grouping)
- Membuat year-to-date (YTD) analyses
- Menghitung usia karyawan atau aset
- Validasi tahun untuk compliance reporting

**Contoh**:
```
CV Komputer Indonesia - Employee Tenure Analysis
- Hire date: 2020-05-15
- Tahun mulai bekerja: =YEAR(B2) → 2020
- Tahun sekarang: =YEAR(TODAY()) → 2026
- Tenure (tahun): =YEAR(TODAY()) - YEAR(B2) → 6 tahun

Interpretasi: Karyawan telah bekerja 6 tahun (sejak 2020)
```

---

#### 1.4 MONTH()
**Definisi**: Mengekstrak bulan dari sebuah tanggal.

**Sintaks**:
```
=MONTH(date)
```

**Parameter**:
- date: Nilai tanggal

**Return Value**: Angka bulan 1-12 (misal: 3 untuk Maret)

**Kegunaan Bisnis**:
- Membuat monthly report dan grouping
- Seasonal analysis (analisis musiman)
- Budget allocation per bulan
- Monthly revenue tracking

**Contoh**:
```
PT Maju Jaya - Monthly Sales Analysis
- Tanggal transaksi: 2026-03-15
- Bulan transaksi: =MONTH(A2) → 3
- Tahun: =YEAR(A2) → 2026
- Periode: "Maret 2026"

Interpretasi: Transaksi terjadi di bulan Maret, dikelompokkan dengan transaksi Maret lainnya
```

---

#### 1.5 DAY()
**Definisi**: Mengekstrak hari (tanggal dalam sebulan) dari sebuah tanggal.

**Sintaks**:
```
=DAY(date)
```

**Parameter**:
- date: Nilai tanggal

**Return Value**: Angka hari 1-31 (misal: 29)

**Kegunaan Bisnis**:
- Membuat daily report dan tracking
- Cut-off date validation
- End-of-month processing
- Payroll scheduling (misal gaji dibayar tanggal 25)

**Contoh**:
```
CV Komputer Indonesia - Payroll Scheduling
- Tanggal hari ini: 2026-03-29
- Hari dalam bulan: =DAY(TODAY()) → 29
- Cek: Apakah sudah melewati tanggal 25? → Ya (gaji sudah bisa diproses)

Interpretasi: Hari ke-29, sudah melewati cut-off payroll tanggal 25
```

---

### Blok 2: Fungsi Penghitung Hari (DATEDIF)

#### 2.1 DATEDIF()
**Definisi**: Menghitung selisih antara dua tanggal dengan berbagai unit (hari, bulan, tahun).

**Sintaks**:
```
=DATEDIF(start_date, end_date, unit)
```

**Parameter**:
- start_date: Tanggal awal
- end_date: Tanggal akhir
- unit: Unit selisih (D=hari, M=bulan, Y=tahun, YM=bulan dalam tahun, MD=hari dalam bulan, YD=hari dalam tahun)

**Return Value**: Angka selisih sesuai unit

**Kegunaan Bisnis**:
- Menghitung masa kerja (tenure) karyawan
- Menghitung usia piutang/receivables aging
- Project duration dan timeline
- Leave/cuti calculation (hari cuti yang digunakan)

**Contoh**:
```
CV Komputer Indonesia - Employee Tenure Report
- Hire date: 2020-05-15
- Tanggal hari ini: 2026-03-29
- Tenure (tahun): =DATEDIF(B2, TODAY(), "Y") → 5 tahun
- Tenure (bulan): =DATEDIF(B2, TODAY(), "YM") → 9 bulan
- Tenure (hari): =DATEDIF(B2, TODAY(), "D") → 2119 hari

Interpretasi: Karyawan telah bekerja 5 tahun 9 bulan (2119 hari)
```

**Unit Penjelasan**:
- "D": Hari (total hari antara dua tanggal)
- "M": Bulan (total bulan)
- "Y": Tahun (total tahun)
- "YM": Bulan tanpa menghitung tahun
- "MD": Hari tanpa menghitung bulan
- "YD": Hari tanpa menghitung tahun

---

### Blok 3: Fungsi Hari Kerja (NETWORKDAYS, WORKDAY)

#### 3.1 NETWORKDAYS()
**Definisi**: Menghitung jumlah hari kerja (Senin-Jumat) antara dua tanggal, excluding weekends dan hari libur.

**Sintaks**:
```
=NETWORKDAYS(start_date, end_date, [holidays])
```

**Parameter**:
- start_date: Tanggal awal
- end_date: Tanggal akhir
- holidays (optional): Range tanggal hari libur yang harus dikecualikan

**Return Value**: Jumlah hari kerja

**Kegunaan Bisnis**:
- Menghitung hari kerja untuk project timeline
- Lead time calculation (waktu pemrosesan)
- Employee leave balance dan working days
- Project scheduling dan resource planning

**Contoh**:
```
CV Komputer Indonesia - Project Timeline
- Start date proyek: 2026-03-01
- End date (target): 2026-03-31
- Hari kerja tersedia: =NETWORKDAYS(B2, B3) → 22 hari kerja
- (Ini exclude Sabtu/Minggu dan hari libur nasional)

Interpretasi: Proyek memiliki 22 hari kerja untuk dikerjakan dalam periode Maret
```

---

#### 3.2 WORKDAY()
**Definisi**: Menghitung tanggal yang merupakan N hari kerja sebelum atau sesudah start date.

**Sintaks**:
```
=WORKDAY(start_date, days, [holidays])
```

**Parameter**:
- start_date: Tanggal awal
- days: Jumlah hari kerja yang akan ditambahkan (negatif untuk mundur)
- holidays (optional): Range hari libur

**Return Value**: Tanggal hasil perhitungan

**Kegunaan Bisnis**:
- Menentukan deadline berdasarkan hari kerja
- Planning jadwal pengiriman (delivery date)
- SLA (Service Level Agreement) calculation
- Project milestone dan task scheduling

**Contoh**:
```
CV Komputer Indonesia - Delivery Schedule
- Order date: 2026-03-01 (Sabtu)
- SLA: Pengiriman dalam 5 hari kerja
- Estimasi pengiriman: =WORKDAY(B2, 5) → 2026-03-06 (Jumat)
- (Skip Sabtu/Minggu, hanya hitung hari kerja)

Interpretasi: Order yang masuk Sabtu akan dikirim Jumat minggu depan (5 hari kerja)
```

---

### Blok 4: Fungsi Manipulasi Tanggal (EDATE)

#### 4.1 EDATE()
**Definisi**: Menambahkan atau mengurangi bulan dari sebuah tanggal.

**Sintaks**:
```
=EDATE(start_date, months)
```

**Parameter**:
- start_date: Tanggal awal
- months: Jumlah bulan yang akan ditambahkan (negatif untuk mundur)

**Return Value**: Tanggal baru setelah penambahan/pengurangan bulan

**Kegunaan Bisnis**:
- Menghitung warranty expiration date
- Contract renewal date planning
- Subscription billing calculation
- Loan/cicilan due date

**Contoh**:
```
CV Komputer Indonesia - Warranty Management
- Tanggal pembelian: 2026-03-15
- Garansi: 24 bulan
- Warranty expiration: =EDATE(B2, 24) → 2028-03-15
- (Tambah 24 bulan dari tanggal pembelian)

Interpretasi: Produk terjamin hingga Maret 2028 (2 tahun dari pembelian)
```

**Perbedaan EDATE vs DATEDIF**:
- EDATE: Menambah/mengurangi bulan (untuk forward/backward planning)
- DATEDIF: Menghitung selisih antara dua tanggal (untuk measurement)

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: Kasus Bisnis CV Komputer Indonesia - HR Department

**Konteks Bisnis**:
CV Komputer Indonesia memiliki departemen HR yang mengelola data karyawan. Data ini mencakup informasi hire date, project timeline, dan warranty management yang memerlukan analisis date-time untuk compliance, performance evaluation, dan resource planning.

### Struktur Kolom & Penjelasan

| No | Kolom | Tipe Data | Keterangan | Contoh |
|----|-------|-----------|-----------|---------|
| A | No | Integer | Nomor urut | 1, 2, 3, ... |
| B | ID Karyawan | Text | Nomor identitas | EMP-001 |
| C | Nama Karyawan | Text | Nama lengkap | Budi Santoso |
| D | Hire Date | Date | Tanggal masuk kerja | 2020-05-15 |
| E | Posisi | Text | Jabatan/posisi | Sales Manager |
| F | Tenure (Tahun) | Integer | Lama bekerja (calculated) | 6 |
| G | Project Start | Date | Tanggal mulai project | 2026-03-01 |
| H | Project End | Date | Tanggal selesai project | 2026-03-31 |
| I | Working Days | Integer | Hari kerja (calculated) | 22 |

**Data untuk Dianalisis**:
- Kolom A-E: Data input (dari sistem HR)
- Kolom F: Hasil calculation DATEDIF/YEAR
- Kolom G-I: Project timeline analysis

### Sample Data Praktik (8 Karyawan)

```
No | ID      | Nama              | Hire Date   | Posisi           | Project Start | Project End
1  | EMP-001 | Budi Santoso      | 2020-05-15  | Sales Manager    | 2026-03-01   | 2026-03-31
2  | EMP-002 | Ani Wijaya        | 2019-08-20  | Finance Manager  | 2026-03-01   | 2026-03-22
3  | EMP-003 | Rudi Firmansah    | 2022-01-10  | Tech Support     | 2026-03-05   | 2026-03-31
4  | EMP-004 | Siti Nurhaliza    | 2021-03-01  | HR Coordinator   | 2026-03-01   | 2026-03-15
5  | EMP-005 | Ahmad Dahlan      | 2023-06-12  | Junior Developer | 2026-03-08   | 2026-03-31
6  | EMP-006 | Dewi Lestari      | 2018-11-05  | Director         | 2026-03-01   | 2026-03-31
7  | EMP-007 | Hendra Wijaya     | 2021-07-22  | Data Analyst     | 2026-03-15   | 2026-03-31
8  | EMP-008 | Linda Putri       | 2022-02-14  | Customer Service | 2026-03-01   | 2026-03-25
```

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Functions
*Fokus: TODAY, YEAR, MONTH, DAY, DATEDIF - Durasi: 40 menit*

#### Latihan 1.1: Hitung Tenure Karyawan dengan YEAR dan DATEDIF
**Waktu**: 12 menit

**Instruksi**:
1. Copy sample data di atas ke Excel (Kolom A sampai E, data 1-8)

2. Di kolom F, buat formula untuk hitung tenure (tahun kerja):
   ```
   Metode 1 (simple):
   =YEAR(TODAY()) - YEAR(D2)
   
   Metode 2 (akurat, menghitung dari hire date):
   =DATEDIF(D2, TODAY(), "Y")
   ```

3. **Analisis**:
   - Siapa karyawan dengan tenure paling lama?
   - Siapa karyawan baru (< 1 tahun)?
   - Rata-rata tenure di perusahaan?

**Pertanyaan**:
- Karyawan mana yang akan mencapai 5 tahun kerja tahun ini?

---

#### Latihan 1.2: Extract Komponen Tanggal (MONTH, YEAR, DAY)
**Waktu**: 8 menit

**Instruksi**:
1. Gunakan data dari Latihan 1.1

2. Buat 3 kolom baru untuk extract:
   ```
   Hire Year: =YEAR(D2)
   Hire Month: =MONTH(D2)
   Hire Day: =DAY(D2)
   ```

3. **Analisis**:
   - Berapa banyak karyawan yang hire bulan sama dengan hari ini (Anniversary month)?
   - Bulan mana yang paling banyak hire?

---

#### Latihan 1.3: Hitung Age of Data dengan TODAY
**Waktu**: 5 menit

**Instruksi**:
1. Gunakan data Latihan 1.1

2. Buat kolom baru:
   ```
   Days Since Hire: =TODAY() - D2
   Months Since Hire: =DATEDIF(D2, TODAY(), "M")
   ```

3. **Interpretasi**:
   - Karyawan mana yang most recently hired?
   - Siapa yang sudah lama tidak ada performance review (estimate dari hire date)?

---

### LEVEL 2: Analisis Lanjutan dengan Date-Time & Project Planning
*Fokus: NETWORKDAYS, WORKDAY, EDATE, Project Timeline - Durasi: 50 menit*

#### Latihan 2.1: Hitung Hari Kerja Project Timeline
**Waktu**: 30 menit

**TAHAP A: Hitung Working Days Menggunakan NETWORKDAYS (10 menit)**

1. Gunakan data dari Latihan 1.1 (Kolom G-H: Project Start dan End)
2. Buat kolom I untuk hitung working days:
   ```
   =NETWORKDAYS(G2, H2)
   ```
   Hasil: Jumlah hari kerja (exclude Sabtu/Minggu)

3. **Verifikasi**:
   - Project mana yang paling lama (banyak working days)?
   - Project mana yang paling singkat?

---

**TAHAP B: Tentukan Deadline Menggunakan WORKDAY (10 menit)**

1. Buat kolom baru "Adjusted Deadline" (jika project harus selesai 15 hari kerja)
2. Formula:
   ```
   =WORKDAY(G2, 15)
   ```
   Hasil: Tanggal yang merupakan 15 hari kerja dari start date

3. **Analisis**:
   - Apakah original deadline (kolom H) tercapai jika ada 15 hari kerja requirement?
   - Mana project yang perlu adjusted deadline?

---

**TAHAP C: Hitung Warranty/Contract Expiration dengan EDATE (10 menit)**

1. Asumsikan project adalah warranty period selama 12 bulan
2. Buat kolom baru "Warranty Expiration":
   ```
   =EDATE(H2, 12)
   ```
   Hasil: Tanggal warranty berakhir (12 bulan setelah project end)

3. **Output**:
   - Setiap project punya warranty period yang jelas
   - Dapat digunakan untuk maintenance planning

---

#### Latihan 2.2: Dashboard HR Analytics dengan Summary Metrics
**Waktu**: 20 menit

**Instruksi**:

1. **Buat Dashboard Area** (di samping data utama):

   ```
   DASHBOARD HR ANALYTICS - MARET 2026
   ===================================
   
   A. EMPLOYEE METRICS:
   - Total Karyawan: =COUNTA(C2:C9)
   - Rata-rata Tenure (Tahun): =AVERAGE(tenure_column)
   - Karyawan Terlama Bekerja: [NAME] - [YEAR] tahun
   - Karyawan Terakhir Hire: [NAME] - [DAYS] hari lalu
   
   B. PROJECT TIMELINE METRICS:
   - Total Working Days Semua Project: =SUM(working_days_column)
   - Rata-rata Working Days per Project: =AVERAGE(working_days_column)
   - Project Dengan Deadline Terdekat: =MIN(H2:H9)
   - Project Dengan Durasi Terlama: [PROJECT] - [DAYS] hari kerja
   
   C. ANNIVERSARY & MILESTONES:
   - Karyawan yang Anniversary Bulan Ini: [LIST]
   - Upcoming 5-Year Milestones: [LIST]
   ```

2. **Gunakan kombinasi formulas**:
   - COUNTA() untuk menghitung total
   - AVERAGE(), MAX(), MIN(), DATEDIF() untuk metrics
   - TODAY() untuk current reference
   - CONCATENATE/& untuk display teks

3. **Interpretasi Dashboard**:
   - Siapa karyawan yang perlu recognition (anniversary)?
   - Project mana yang paling kritis (banyak working days)?
   - Siapa yang akan 5 tahun tenure dalam 6 bulan ke depan?

---

## 📋 INSTRUKSI PENGUMPULAN & EVALUASI

### Checklist Pengerjaan Praktikum (Durasi: 90 Menit)

**LEVEL 1 - Penerapan Dasar (40 menit)**:
- [ ] Latihan 1.1: YEAR, DATEDIF untuk tenure calculation (12 menit)
- [ ] Latihan 1.2: MONTH, YEAR, DAY untuk extract komponen (8 menit)
- [ ] Latihan 1.3: TODAY, DATEDIF untuk age of data (5 menit)
- [ ] Buffer/Review: 15 menit

**LEVEL 2 - Analisis Lanjutan (50 menit)**:
- [ ] Latihan 2.1 TAHAP A: NETWORKDAYS untuk hari kerja (10 menit)
- [ ] Latihan 2.1 TAHAP B: WORKDAY untuk deadline planning (10 menit)
- [ ] Latihan 2.1 TAHAP C: EDATE untuk warranty/contract (10 menit)
- [ ] Latihan 2.2: Dashboard HR Analytics dengan metrics (20 menit)

**Finalisasi (10 menit min)**:
- [ ] File saved dengan naming: `NAMA_NIM_Week4_DateTimeFunction.xlsx`
- [ ] Setiap worksheet diberi title dan keterangan kolom
- [ ] Formulas ter-dokumentasi (keterangan di samping/bawah)

---

### Rubric Penilaian (Outcome Based Education)

| Learning Outcome | Indikator Ketercapaian | Skor |
|-----------------|-------------------------|------|
| **C3 (Apply)** - Menerapkan Date-Time functions | Semua formula TODAY, NOW, YEAR, MONTH, DAY, DATEDIF, NETWORKDAYS, WORKDAY, EDATE diinput dengan benar; output sesuai harapan; tanpa error | 50% |
| **C4 (Analyze)** - Menganalisis data temporal & project timeline | Dashboard dibuat dengan HR metrics yang tepat; timeline planning akurat; interpretasi business insight jelas; milestone/deadline planning informatif | 50% |

**Kriteria Passing**: Score ≥ 70 (C3 ≥ 35 + C4 ≥ 35)



---

## 📚 TIPS & TROUBLESHOOTING PRAKTIKUM

### Common Mistakes & Fixes

| Masalah | Solusi |
|---------|--------|
| DATEDIF error #NAME? | Fungsi DATEDIF mungkin tidak tersedia di versi Excel lama. Gunakan alternatif: =INT(end_date - start_date) atau =(YEAR(end_date)-YEAR(start_date))*365 + (MONTH(end_date)-MONTH(start_date))*30 + (DAY(end_date)-DAY(start_date)) |
| TODAY() tidak berubah otomatis | TODAY() update hanya saat file dibuka ulang atau dicalculate. Tekan Ctrl+Shift+F9 untuk force recalculate. |
| YEAR/MONTH/DAY menghasilkan 0 | Pastikan input adalah format Date, bukan Text. Format sel mungkin Text - ubah ke Date format. |
| NETWORKDAYS error #NUM! | Range holidays mungkin overlap atau format salah. Pastikan holiday dates valid dan tidak lebih besar dari end_date. |
| WORKDAY menghasilkan tanggal salah | Pastikan days parameter adalah angka, bukan format yang salah. Days negatif untuk mundur (misal -5 untuk 5 hari kerja sebelum). |
| EDATE hasil tidak presisi | EDATE menambah exact months, bukan days. Jika tanggal input 1/31 + 1 month = 2/28 (karena Feb tidak ada 31). |

### Shortcut & Tips Efisiensi

1. **Quick Date Format**: Gunakan Ctrl+Shift+3 (Windows) atau Cmd+Shift+3 (Mac) untuk quick date format.
2. **Auto Date Entry**: Type tanggal alami (misal "1 Mar 2026") dan Excel akan auto-recognize format.
3. **Date Series**: Select tanggal cell, drag fill handle dengan pattern untuk auto-fill date series.
4. **TODAY vs Static Date**: Gunakan TODAY() untuk tanggal yang selalu current; gunakan fixed date untuk historical analysis.
5. **DATEDIF Unit Combinations**: Combine units untuk hasil yang fleksibel (misal "Y" untuk year, "M" untuk month total, "YM" untuk month only).

---

## 🔧 INSTRUKSI UNTUK AGENT: HTML GENERATION

### Purpose
Markdown ini adalah blueprint untuk generate HTML file yang professional, accessible, dan sesuai Tailwind CSS styling.

### Key Points untuk Agent

1. **Layout & Structure**
   - Gunakan 3-column grid (sidebar + main content) seperti week1.html, week2.html, & week3.html
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
   - Progress tracker (visual checklist dengan input type="checkbox")
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
- [ ] **Extraction/Manipulation**: Text/Date/Number processing lessons
- [ ] **Documentation**: Jelas, professional, mudah dipahami
- [ ] **Rubric Penilaian**: OBE-aligned dengan C3 (Apply) + C4 (Analyze)

---

**End of Document**

*Dokumen ini adalah template comprehensive untuk latihan Excel Date-Time Functions. 
Dipersiapkan untuk Outcome Based Education, business-context relevant, 
dan AI-friendly untuk code generation.*
