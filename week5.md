# Latihan Pertemuan 5: Conditional Logic Functions (IF) Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-005
- **Pertemuan**: 5
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Pemula (Beginner)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 (Basic), Week 2 (Formula), Week 3 (Mathematical & Text Functions), dan Week 4 (Date-Time Functions)

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi logika kondisional dalam konteks bisnis nyata
   - Menggunakan IF() untuk membuat keputusan berbasis kondisi tunggal
   - Mengimplementasikan nested IF() untuk multiple conditions
   - Mengaplikasikan IFS() untuk simplifikasi multi-kondisi
   - Menggunakan AND(), OR(), NOT() dalam kombinasi dengan IF
   - Membuat formula logika untuk validasi dan klasifikasi data

2. **C4 (Analyze)** - Menganalisis data dengan logika kondisional untuk insight bisnis
   - Membuat sales performance classification berdasarkan threshold
   - Menganalisis customer tier dan segmentation
   - Merancang automated scoring dan eligibility system
   - Membuat if-then-else decision logic untuk business rules
   - Menginterpretasikan hasil kondisional untuk strategic planning

### Notifikasi Kompatibilitas Excel
- **IFS()** tidak tersedia di Excel 2013 dan Excel 2016 (tersedia mulai Excel 2019 / Microsoft 365).
- Materi IFS tetap diperkenalkan sebagai wawasan tambahan.
- **Latihan IFS dikerjakan mandiri di luar kelas/perkuliahan** dan tidak masuk checklist penilaian utama.

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: Fungsi Logika Dasar (IF)

#### 1.1 IF() - Kondisi Tunggal
**Definisi**: Mengembalikan satu nilai jika kondisi TRUE, nilai lain jika FALSE.

**Sintaks**:
```
=IF(condition, value_if_true, value_if_false)
```

**Parameter**:
- condition: Pernyataan logika yang dievaluasi (TRUE/FALSE)
- value_if_true: Nilai yang dikembalikan jika kondisi TRUE
- value_if_false: Nilai yang dikembalikan jika kondisi FALSE

**Return Value**: value_if_true atau value_if_false (tipe bergantung input)

**Kegunaan Bisnis**:
- Menentukan bonus berdasarkan target tercapai
- Validasi data (pass/fail)
- Kategori status (aktif/nonaktif)
- Membership tier classification (premium/regular)

**Contoh**:
```
CV Komputer Indonesia - Sales Performance
- Penjualan karyawan: 50,000,000 (target: 40,000,000)
- Formula: =IF(B2 >= 40000000, "Target Tercapai", "Target Tidak Tercapai")
- Hasil: "Target Tercapai"

Interpretasi: Penjualan melebihi target, karyawan mendapat reward
```

**Karakteristik**:
- Evaluasi TRUE/FALSE hanya satu kali
- value_if_false bersifat optional (default: FALSE jika tidak ada)
- Bisa nested untuk multi-kondisi

---

#### 1.2 IF() - Nested (Kondisi Berganda)
**Definisi**: Penggunaan IF di dalam IF untuk menangani multiple conditions.

**Sintaks**:
```
=IF(condition1, value_if_true1, IF(condition2, value_if_true2, value_if_false2))
```

**Parameter**:
- Struktur bersarang dengan kondisi bertingkat

**Return Value**: Hasil dari kondisi TRUE pertama yang ditemukan

**Kegunaan Bisnis**:
- Multi-level categorization (Bronze/Silver/Gold/Platinum)
- Grading system (A/B/C/D/E)
- Commission/bonus tiers
- Risk assessment levels

**Contoh**:
```
CV Komputer Indonesia - Sales Commission Structure
- Penjualan: 75,000,000
- Formula:
  =IF(B2 >= 100000000, 15%, IF(B2 >= 75000000, 12%, IF(B2 >= 50000000, 10%, 5%)))
- Hasil: 12% (karena penjualan 75jt masuk tier >= 75jt)

Interpretasi: Komisi 12% = 9,000,000
```

---

### Blok 2: Logical Operators (AND, OR, NOT)

#### 2.1 AND() - Semua Kondisi Harus TRUE
**Definisi**: Mengembalikan TRUE hanya jika semua kondisi TRUE.

**Sintaks**:
```
=AND(condition1, condition2, condition3, ...)
```

**Parameter**:
- condition1, condition2, ...: Kondisi yang dievaluasi

**Return Value**: TRUE atau FALSE

**Kegunaan Bisnis**:
- Validasi multiple requirements (semua harus terpenuhi)
- Eligibility checking (layak/tidak layak kredit)
- Quality control (semua parameter OK)

**Contoh**:
```
CV Komputer Indonesia - Promo Eligibility
- Kondisi: Pembelian >= 50jt AND member aktif AND tanpa hutang
- Formula: =AND(B2 >= 50000000, C2 = "Aktif", D2 = "Lunas")
- Hasil: TRUE (semua kondisi terpenuhi)

Interpretasi: Pelanggan eligible untuk promo diskon 20%
```

---

#### 2.2 OR() - Minimal Satu Kondisi TRUE
**Definisi**: Mengembalikan TRUE jika minimal satu kondisi TRUE.

**Sintaks**:
```
=OR(condition1, condition2, condition3, ...)
```

**Parameter**:
- condition1, condition2, ...: Kondisi yang dievaluasi

**Return Value**: TRUE atau FALSE

**Kegunaan Bisnis**:
- Alert system (jika salah satu threshold terlampaui)
- Exception handling (apply rule jika salah satu kondisi ada)
- Risk flagging (red flag jika ada warning sign)

**Contoh**:
```
CV Komputer Indonesia - Risk Assessment
- Kondisi: Penjualan turun > 20% OR Status kredit merah OR Product return > 10%
- Formula: =OR(E2 < -0.2, F2 = "Merah", G2 > 0.1)
- Hasil: TRUE (penjualan turun, flagged untuk follow-up)

Interpretasi: Perlu tindakan follow-up untuk pelanggan ini
```

---

#### 2.3 NOT() - Inversi Kondisi
**Definisi**: Mengembalikan kebalikan dari nilai logika (TRUE → FALSE, FALSE → TRUE).

**Sintaks**:
```
=NOT(condition)
```

**Parameter**:
- condition: Pernyataan logika

**Return Value**: Kebalikan dari condition

**Kegunaan Bisnis**:
- Exemption logic (jika BUKAN premium member)
- Negative conditions (jika BUKAN complete)
- Inversion untuk cleaner formula

**Contoh**:
```
CV Komputer Indonesia - Shipping Cost
- Jika BUKAN pembelian online, tambah biaya handling 50,000
- Formula: =IF(NOT(B2 = "Online"), A2 + 50000, A2)
- Hasil: Biaya disesuaikan berdasarkan channel

Interpretasi: Biaya handling hanya untuk non-online orders
```

---

### Blok 3: IFS() - Multi-Condition Simplified

**Catatan Kompatibilitas**:
- IFS tersedia di Excel 2019 / Microsoft 365.
- Untuk Excel 2013/2016, gunakan nested IF sebagai alternatif.
- Blok ini bersifat latihan mandiri (opsional) di luar kelas.

#### 3.1 IFS() - Multiple IF-Else Conditions
**Definisi**: Mengevaluasi multiple conditions tanpa nested IF, return nilai pertama kondisi TRUE.

**Sintaks**:
```
=IFS(condition1, value1, condition2, value2, condition3, value3, ...)
```

**Parameter**:
- condition1, condition2, ...: Kondisi dalam urutan evaluasi
- value1, value2, ...: Nilai yang dikembalikan jika kondisi TRUE

**Return Value**: Nilai dari kondisi TRUE pertama

**Kegunaan Bisnis**:
- Clean multi-tier classification
- Simplified nested IF (lebih readable)
- Grade/score mapping

**Contoh**:
```
CV Komputer Indonesia - Customer Segmentation
- Revenue: 150,000,000
- Formula:
  =IFS(B2 >= 200000000, "Platinum", B2 >= 100000000, "Gold", 
       B2 >= 50000000, "Silver", B2 >= 0, "Bronze")
- Hasil: "Gold" (150jt masuk tier Platinum? No → Gold tier)

Interpretasi: Customer masuk segmen Gold untuk special pricing
```

---

### Blok 4: Complex Conditions - Kombinasi Operator

#### 4.1 IF + AND - Multiple Kondisi All-Must-TRUE
**Definisi**: Kombinasi IF dan AND untuk validasi multiple conditions harus semua TRUE.

**Sintaks**:
```
=IF(AND(condition1, condition2, condition3), value_true, value_false)
```

**Kegunaan Bisnis**:
- Bonus calculation (penjualan target AND customer satisfaction > 4 AND attendance OK)
- Promotion eligibility (purchase >= threshold AND member AND payment on-time)
- Access control (role = admin AND date within period AND IP valid)

**Contoh**:
```
CV Komputer Indonesia - Bonus Eligibility
- Kondisi: Penjualan >= 40jt AND Customer satisfaction >= 4.5 AND Attendance = Lengkap
- Formula:
  =IF(AND(B2 >= 40000000, C2 >= 4.5, D2 = "Lengkap"), "Eligible", "Not Eligible")
- Hasil: "Eligible" (semua kondisi terpenuhi)

Interpretasi: Karyawan layak dapat bonus bulanan
```

---

#### 4.2 IF + OR - Multiple Kondisi Any-Can-TRUE
**Definisi**: Kombinasi IF dan OR untuk trigger action jika salah satu kondisi TRUE.

**Sintaks**:
```
=IF(OR(condition1, condition2, condition3), value_true, value_false)
```

**Kegunaan Bisnis**:
- Exception handling (jika ada error di salah satu field)
- Alert system (jika stock kurang OR expiry date dekat OR quality issue)
- VIP treatment (jika high-value customer OR loyalty member OR executive)

**Contoh**:
```
CV Komputer Indonesia - Urgent Handling
- Kondisi: Stock < 5 OR Expiry dalam 30 hari OR Quality issue
- Formula:
  =IF(OR(B2 < 5, C2 < 30, D2 = "Ada"), "URGENT", "Regular")
- Hasil: "URGENT" (ada salah satu kondisi urgent)

Interpretasi: Produk perlu immediate action
```

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: Kasus Bisnis CV Komputer Indonesia - Sales & Customer Performance

**Konteks Bisnis**:
CV Komputer Indonesia memiliki departemen penjualan yang mengelola data transaksi. Data ini mencakup informasi penjualan, customer status, performance metrics yang memerlukan evaluasi berbasis kondisi untuk bonus determination, customer segmentation, dan business intelligence.

### Struktur Kolom & Penjelasan

| No | Kolom | Tipe Data | Keterangan | Contoh |
|----|-------|-----------|-----------|---------|
| A | No | Integer | Nomor urut | 1, 2, 3, ... |
| B | Sales ID | Text | ID Sales Person | SAL-001 |
| C | Sales Name | Text | Nama Sales | Budi Santoso |
| D | Monthly Sales | Currency | Total penjualan bulan (Rp) | 65,000,000 |
| E | Customer Satisfaction | Decimal | Rating 1-5 | 4.2 |
| F | Attendance | Text | Status kehadiran | Lengkap |
| G | Payment Status | Text | Status pembayaran | Lunas |
| H | Membership Tier | Text | Tier customer (calculated) | Gold |
| I | Bonus Eligible | Text | Eligible/Not Eligible (calculated) | Eligible |
| J | Action Required | Text | Follow-up action (calculated) | None |

**Data untuk Dianalisis**:
- Kolom A-G: Data input (dari sistem sales)
- Kolom H-J: Hasil calculation IF/IFS/AND/OR

### Sample Data Praktik (8 Sales Person)

```
No | Sales ID | Sales Name        | Monthly Sales | Cust Satisfaction | Attendance | Payment Status | Membership Tier | Bonus Eligible | Action Required
1  | SAL-001  | Budi Santoso      | 65,000,000    | 4.2               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
2  | SAL-002  | Ani Wijaya        | 95,000,000    | 4.8               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
3  | SAL-003  | Rudi Firmansah    | 35,000,000    | 3.5               | Tidak      | Hutang         | [Formula]       | [Formula]      | [Formula]
4  | SAL-004  | Siti Nurhaliza    | 120,000,000   | 4.6               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
5  | SAL-005  | Ahmad Dahlan      | 48,000,000    | 4.0               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
6  | SAL-006  | Dewi Lestari      | 180,000,000   | 4.9               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
7  | SAL-007  | Hendra Wijaya     | 42,000,000    | 3.8               | Tidak      | Lunas          | [Formula]       | [Formula]      | [Formula]
8  | SAL-008  | Linda Putri       | 155,000,000   | 4.7               | Lengkap    | Lunas          | [Formula]       | [Formula]      | [Formula]
```

**Keterangan Data**:
- Monthly Sales: Target penjualan 50,000,000
- Customer Satisfaction: Skala 1-5, target >= 4.5 untuk bonus
- Attendance: Lengkap = hadir, Tidak = absence/cuti
- Payment Status: Lunas = no debt, Hutang = outstanding debt
- Membership Tier: Platinum (>= 150jt), Gold (100-149jt), Silver (50-99jt), Bronze (< 50jt)

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Functions
*Fokus: IF, Nested IF, AND, OR - Durasi: 40 menit*

#### Latihan 1.1 (Mandiri - Opsional): Klasifikasi Membership Tier dengan IFS
**Waktu**: 12 menit

**Keterangan Versi**:
- Excel 2019 / Microsoft 365: gunakan IFS sesuai instruksi.
- Excel 2013 / 2016: gunakan nested IF (alternatif), latihan ini tidak dinilai pada checklist kelas.

**Instruksi**:
1. Copy sample data di atas ke Excel (Kolom A sampai G, data 1-8)

2. Di kolom H, buat formula untuk klasifikasi membership tier berdasarkan penjualan:
   ```
   Platinum: >= 150,000,000
   Gold: 100,000,000 - 149,999,999
   Silver: 50,000,000 - 99,999,999
   Bronze: < 50,000,000
   
   Formula:
   =IFS(D2 >= 150000000, "Platinum", D2 >= 100000000, "Gold", 
        D2 >= 50000000, "Silver", D2 >= 0, "Bronze")
   ```

3. **Analisis**:
   - Berapa sales yang tier Platinum?
   - Berapa sales yang tier Silver?
   - Rata-rata penjualan per tier?

**Pertanyaan**:
- Tier mana yang paling banyak? Interpretasi untuk strategi?

---

#### Latihan 1.2: Eligibility Check dengan IF + AND
**Waktu**: 15 menit

**Instruksi**:
1. Gunakan data dari Latihan 1.1 + tambahan kolom Membership Tier

2. Di kolom I, buat formula untuk check bonus eligibility:
   ```
   Kriteria bonus:
   - Monthly Sales >= 40,000,000
   - Customer Satisfaction >= 4.0
   - Attendance = "Lengkap"
   - Payment Status = "Lunas"
   
   Formula (menggunakan nested IF + AND):
   =IF(AND(D2 >= 40000000, E2 >= 4.0, F2 = "Lengkap", G2 = "Lunas"), 
        "Eligible", 
        "Not Eligible")
   ```

3. **Analisis**:
   - Berapa sales yang eligible bonus?
   - Siapa yang tidak eligible? Alasan kenapa?
   - Apakah ada correlation antara tier dan bonus eligibility?

---

#### Latihan 1.3: Simple Flag dengan IF + Nested
**Waktu**: 13 menit

**Instruksi**:
1. Gunakan data dari Latihan 1.1 dan 1.2

2. Di kolom J, buat formula untuk flagging action required:
   ```
   Logika:
   - Jika penjualan < 40jt AND attendance = "Tidak"  → "Priority Review"
   - Jika payment = "Hutang"                           → "Payment Follow-up"
   - Jika satisfaction < 3.5                           → "Coaching Required"
   - Jika tidak ada issue                              → "None"
   
   Formula (nested IF + OR):
   =IF(OR(AND(D2 < 40000000, F2 = "Tidak"), G2 = "Hutang", E2 < 3.5),
        IF(AND(D2 < 40000000, F2 = "Tidak"), "Priority Review",
           IF(G2 = "Hutang", "Payment Follow-up",
              "Coaching Required")),
        "None")
   ```

3. **Analisis**:
   - Siapa yang butuh coaching?
   - Siapa yang butuh payment follow-up?
   - Siapa yang priority review?

---

### LEVEL 2: Analisis Lanjut & Business Intelligence
*Fokus: Complex Conditions, Dashboard, Strategic Insights - Durasi: 50 menit*

#### Latihan 2.1: Multi-Tier Commission Structure
**Waktu**: 15 menit

**Instruksi**:
1. Buat kolom "Commission (%)" di kolom baru (K)

2. Buat formula commission berdasarkan tier berganda:
   ```
   Basis: Penjualan + Customer Satisfaction + Bonus Status
   
   Tier Commission:
   - Platinum + Eligible        → 15%
   - Platinum + Not Eligible    → 10%
   - Gold + Eligible            → 12%
   - Gold + Not Eligible        → 8%
   - Silver + Eligible          → 10%
   - Silver + Not Eligible      → 6%
   - Bronze + Eligible          → 8%
   - Bronze + Not Eligible      → 4%
   
   Formula menggunakan nested IF + referensi H & I:
   =IF(H2 = "Platinum", 
        IF(I2 = "Eligible", 0.15, 0.10),
        IF(H2 = "Gold",
           IF(I2 = "Eligible", 0.12, 0.08),
           IF(H2 = "Silver",
              IF(I2 = "Eligible", 0.10, 0.06),
              IF(I2 = "Eligible", 0.08, 0.04))))
   ```

3. Buat kolom "Commission Amount (Rp)" di kolom L:
   ```
   =D2 * K2
   ```

4. **Analisis**:
   - Total commission yang keluar berapa?
   - Sales dengan commission tertinggi?
   - Rata-rata commission rate per tier?

---

#### Latihan 2.2: Warranty & Risk Flagging System
**Waktu**: 15 menit

**Instruksi**:
1. Tambah kolom "Risk Level" di kolom M

2. Buat risk assessment logic:
   ```
   Kondisi Risk:
   - RED Alert: Penjualan < 40jt OR Satisfaction < 3.5 OR Payment = "Hutang"
   - YELLOW Alert: Penjualan < 50jt OR Attendance = "Tidak" OR Satisfaction < 4.0
   - GREEN (Normal): Lainnya
   
   Formula:
   =IF(OR(D2 < 40000000, E2 < 3.5, G2 = "Hutang"),
        "RED",
        IF(OR(D2 < 50000000, F2 = "Tidak", E2 < 4.0),
           "YELLOW",
           "GREEN"))
   ```

3. **Analisis**:
   - Berapa sales dengan RED risk? Prioritas follow-up?
   - Berapa sales dengan YELLOW? Coaching program?
   - Trend risk level? Business impact?

---

#### Latihan 2.3: Sales Performance Dashboard
**Waktu**: 20 menit

**Instruksi**:
1. **Buat Dashboard Area dengan metrics:**
   ```
   A. SALES METRICS
   - Total Sales: =SUM(D2:D9)
   - Avg Sales: =AVERAGE(D2:D9)
   - Target Achievers (>= 50jt): 
     =COUNTIF(D2:D9, ">="&50000000)
   - % Target Achievement:
     =COUNTIF(D2:D9, ">="&50000000) / COUNTA(D2:D9)
   
   B. QUALITY METRICS
   - Avg Satisfaction: =AVERAGE(E2:E9)
   - Full Attendance Count: =COUNTIF(F2:F9, "Lengkap")
   - Full Payment Count: =COUNTIF(G2:G9, "Lunas")
   
   C. BONUS & COMMISSION
   - Eligible for Bonus: =COUNTIF(I2:I9, "Eligible")
   - Total Commission: =SUM(L2:L9)
   - Avg Commission per Person: =AVERAGE(L2:L9)
   
   D. RISK ASSESSMENT
   - RED Alert Count: =COUNTIF(M2:M9, "RED")
   - YELLOW Alert Count: =COUNTIF(M2:M9, "YELLOW")
   - GREEN (Normal) Count: =COUNTIF(M2:M9, "GREEN")
   ```

2. **Buat Dashboard Table** dengan 4 kolom (Metrik, Value, Target, Status):
   ```
   Contoh:
   Metrik | Value | Target | Status (IF pairing)
   Avg Sales | [=AVERAGE(...)] | 75,000,000 | =IF([value]>=[target], "ON TRACK", "BELOW TARGET")
   Avg Satisfaction | [=AVERAGE(...)] | 4.0 | =IF([value]>=[target], "GOOD", "NEED IMPROVEMENT")
   ```

3. **Analisis**:
   - Sales team performance overall?
   - Biggest bottleneck (sales, quality, atau compliance)?
   - Top performers? Bottom performers?
   - Rekomendasi action plan?

---

## 📋 INSTRUKSI PENGUMPULAN

### Format File
- Nama file: **NAMA_NIM_Week5_ConditionalLogic.xlsx**
- Sheet utama: Data dan formula
- Sheet dashboard: Dashboard metrics (optional tapi recommended)
- Semua formula visible (tidak di-hide)

### Checklist Pengerjaan (90 menit):
- [ ] LEVEL 1 (40 min): Latihan 1.2 dan 1.3 selesai tanpa error
- [ ] LEVEL 2 (50 min): Latihan 2.1, 2.2, 2.3 selesai + dashboard complete
- [ ] Latihan 1.1 (IFS) dikerjakan mandiri di luar kelas/perkuliahan (opsional, tidak dinilai)
- [ ] Dashboard: Metrics calculation dan insights summary
- [ ] File save dengan naming sesuai ketentuan

### Rubric Penilaian:

| Aspek | Kriteria | Bobot |
|-------|----------|-------|
| **C3 (Apply)** | Semua formula IF, Nested IF, AND, OR diinput benar; logic sesuai business rule; hasil calculation akurat | 50% |
| **C4 (Analyze)** | Dashboard metrics accurate; insights jelas; action recommendation berdasarkan data; interpretasi business context tepat | 50% |

**Passing**: Score ≥ 70 (C3 ≥ 35 + C4 ≥ 35)

---

## 💡 TIPS & TROUBLESHOOTING

### Common Mistakes:

1. **IF Syntax Error**:
   ```
   ❌ SALAH: =IF(D2>=40000000, "OK")  [missing value_if_false]
   ✅ BENAR: =IF(D2>=40000000, "OK", "NOT OK")
   ```

2. **Nested IF - Kurung Terlalu Banyak**:
   ```
   ❌ SALAH: =IF(A, B, IF(C, D, E)))) [kurung tidak seimbang]
   ✅ BENAR: =IF(A, B, IF(C, D, E))
   ```

3. **AND/OR Logic Error**:
   ```
   ❌ SALAH: =AND(D2 >= 40000000, D2 <= 150000000)  [semua harus true]
   ✅ BENAR untuk range: =OR(D2 < 40000000, D2 > 150000000)  [jika ada error flag]
   ```

4. **IFS Evaluation Order**:
   ```
   ❌ SALAH: =IFS(D2 >= 50jt, "Silver", D2 >= 100jt, "Gold", D2 >= 150jt, "Platinum")
        [Gold dan Platinum tidak akan dicapai karena 100jt dan 150jt sudah >= 50jt]
   ✅ BENAR: =IFS(D2 >= 150jt, "Platinum", D2 >= 100jt, "Gold", D2 >= 50jt, "Silver", TRUE, "Bronze")
        [urutan descending, guaranteed match]
   ```

5. **Data Type Mismatch**:
   ```
   ❌ SALAH: =IF("Lunas" = G2, ...)  [quotes pada satu pihak]
   ✅ BENAR: =IF(G2 = "Lunas", ...)  [konsisten]
   ```

### Troubleshooting Tips:

- **Formula shows #NAME?**: Check typo di function name (IF vs IFF)
- **Formula shows #VALUE?**: Data type mismatch atau invalid operator
- **Logic tidak sesuai expectation**: Trace langkah-per-langkah dengan nested IF sederhana dulu
- **IFS tidak match**: Pastikan urutan kondisi dari most specific ke least specific

### Function Alternatives:

- **Nested IF** → **IFS** (cleaner, lebih readable)
- **IF + OR** → **Switch** (Excel 365, untuk exact matching)
- **Multiple IF** → **Lookup functions** (VLOOKUP jika ada lookup table)

### Catatan Kompatibilitas Versi:

- **IFS**: tersedia mulai Excel 2019 / Microsoft 365.
- Untuk Excel 2013/2016, gunakan nested IF.
- Untuk modul ini, latihan IFS tidak masuk checklist penilaian kelas.

---

## 📖 INSTRUKSI UNTUK AGENT / TEACHING ASSISTANT

### Learning Objectives (Detailed):

1. **IF Mastery**:
   - Students memahami IF sebagai fundamental decision-making tool
   - Comfortable dengan nested IF untuk ≤ 3-4 levels
   - Know when to use IF vs IFS (readability threshold)

2. **Logical Operators**:
   - AND: All conditions must be TRUE (conjunction)
   - OR: At least one condition must be TRUE (disjunction)
   - NOT: Invert logic
   - Complex combinations: IF(AND(...), IF(OR(...), ...))

3. **Business Application**:
   - Translate business rules → Excel formulas
   - Recognize patterns: eligibility, classification, flagging
   - Create decision logic for automation

### Common Misconceptions to Clarify:

| Misconception | Correction |
|---|---|
| IF returns only text | IF dapat return number, text, date, atau formula result |
| Nested IF bisa unlimited | Best practice: max 3-4 levels (readability) |
| AND/OR hanya dengan IF | AND/OR berdiri sendiri return TRUE/FALSE, combine dengan IF untuk action |
| IFS lebih efficient | IFS cleaner tapi tidak lebih "efficient" - sama evalnya |
| NOT selalu diperlukan | NOT optional - gunakan saat elegant (negation logic) |

### Assessment Tips:

- **Check**: Apakah formula match dengan business requirement?
- **Trace**: Untuk complex nested, trace dengan example value
- **Validate**: Test dengan edge cases (boundary values)
- **Optimize**: Suggest simplified version jika ada redundancy

### Extensions for Advanced Students:

1. **Error Handling**: `=IFERROR(IF(...), "ERROR")`
2. **Conditional Formatting** + IF formula untuk visual dashboard
3. **Array Formula**: `=IF(range>=threshold, "yes", "no")` untuk bulk evaluation
4. **VBA/Macro**: Conditional logic untuk automation

### Kebijakan Modul Berikutnya (MD Berikutnya):

- Jika ada function yang **tidak tersedia** di Excel 2013/2016, selalu beri notifikasi kompatibilitas.
- Function yang tidak tersedia di Excel 2013/2016 **tidak dimasukkan ke checklist penilaian utama**.
- Function tersebut tetap boleh diberikan sebagai **latihan mandiri** untuk eksplorasi siswa di luar kelas.

---

## 📋 TEMPLATE CHECKLIST UNTUK COPY-PASTE

Mahasiswa dapat menggunakan template ini untuk track progress:

```
WEEK 5 - CONDITIONAL LOGIC CHECKLIST

LEVEL 1 (Target: selesai 40 menit):
□ 1.2 - IF+AND Bonus Eligibility (15 min)
  ├─ Formula input di kolom I
  ├─ AND operator dengan 4 kondisi
  └─ Analisis: 3 critical questions

□ 1.3 - Nested IF Action Flagging (13 min)
  ├─ Formula input di kolom J
  ├─ Nested IF 3 levels
  └─ Analisis: Action items identified

LEVEL 2 (Target: selesai 50 menit):
□ 2.1 - Multi-Tier Commission (15 min)
  ├─ Commission % di kolom K
  ├─ Commission Amount di kolom L
  ├─ Nested IF 4 levels
  └─ Analisis: 3 business metrics

□ 2.2 - Risk Flagging System (15 min)
  ├─ Risk Level di kolom M
  ├─ RED/YELLOW/GREEN classification
  ├─ IF + OR combination
  └─ Analisis: Risk distribution

□ 2.3 - Dashboard (20 min)
  ├─ Sales Metrics (4 KPIs)
  ├─ Quality Metrics (3 KPIs)
  ├─ Bonus & Commission (3 KPIs)
  ├─ Risk Assessment (3 KPIs)
  ├─ Summary Table (Metrik + Target + Status)
  └─ Interpretasi: Top 3 insights

LATIHAN MANDIRI (DI LUAR KELAS - OPSIONAL):
□ 1.1 - IFS Membership Tier (12 min)
   ├─ Khusus Excel 2019 / Microsoft 365
   ├─ Alternatif Excel 2013/2016: Nested IF
   └─ Tidak masuk checklist penilaian utama

FINALISASI:
□ File naming: NAMA_NIM_Week5_ConditionalLogic.xlsx
□ Semua formula visible (tidak hidden)
□ Sheet terorganisir (Data + Dashboard)
□ No error values (#NAME?, #VALUE?, #DIV/0!)

Total Time: 90 menit
Passing Score: ≥ 70 (C3 ≥ 35 + C4 ≥ 35)
```

---

## End of Document
