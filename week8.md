# Latihan Pertemuan 8: Financial Functions Microsoft Excel

## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
- **Kode Modul**: EXCEL-008
- **Pertemuan**: 8
- **Durasi Praktik**: 90 menit
- **Tingkat Kesulitan**: Menengah (Intermediate)
- **Prasyarat**: Mahasiswa telah menguasai Week 1 s.d. Week 7

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:

1. **C3 (Apply)** - Menerapkan fungsi finansial dalam konteks bisnis nyata
   - Menggunakan PMT() untuk menghitung angsuran pinjaman periodik
   - Menggunakan PV() dan NPER() untuk analisis skenario pembiayaan
   - Menggunakan IPMT() dan PPMT() untuk memahami komposisi pembayaran (bunga vs pokok)
   - Membuat tabel amortisasi (riwayat pembayaran) dari awal sampai akhir tenor

2. **C4 (Analyze)** - Menganalisis skema pembiayaan dengan insight bisnis
   - Menganalisis dampak suku bunga terhadap total pembayaran dan lamanya tenor
   - Membandingkan skenario pembiayaan berbeda (tenor panjang vs pendek, rate berbeda)
   - Merumuskan rekomendasi pembiayaan optimal berdasarkan cash flow bisnis

### Apa Itu OBE (Outcome Based Education)?
OBE adalah pendekatan pembelajaran yang fokus pada **hasil akhir yang terukur** (learning outcomes) bukan hanya pada proses. Setiap learning outcome dirancang menggunakan **Bloom's Taxonomy** (C1=Ingat, C2=Pahami, C3=Aplikasi, C4=Analisis, C5=Evaluasi, C6=Kreasi). 
- Level C3 (Apply) berarti mahasiswa bisa **menggunakan** fungsi finansial pada data nyata.
- Level C4 (Analyze) berarti mahasiswa bisa **membandingkan dan menginterpretasikan** hasil analisis.

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: Fungsi PMT() - Payment (Angsuran Tetap)

#### 1.1 PMT() Dasar
**Definisi**: Menghitung pembayaran periodik yang sama (angsuran tetap) untuk melunasi pinjaman atau mencapai target investasi.

**Sintaks**:
```
=PMT(rate, nper, pv, [fv], [type])
```

**Parameter**:
- rate: Suku bunga per periode (contoh: 9%/12 untuk bulanan)
- nper: Jumlah periode pembayaran (contoh: 36 bulan)
- pv: Nilai sekarang (present value = jumlah pinjaman). **Gunakan tanda negatif** agar hasil positif.
- fv (opsional): Nilai masa depan (biasanya 0 untuk pinjaman penuh)
- type (opsional): 0=pembayaran di akhir periode (default), 1=pembayaran di awal periode

**Return Value**: Besarnya angsuran per periode (nilai tunggal dalam Rp)

**Kegunaan Bisnis**:
- Menghitung angsuran kredit mobil/rumah/perlengkapan
- Perencanaan cicilan untuk pembelian aset
- Estimasi kas keluar periodik untuk perencanaan arus kas
- Evaluasi kelayakan pembiayaan berdasarkan kapasitas angsuran

**Contoh**:
```
PT Maju Jaya - Pembiayaan Perangkat Kantor
- Jumlah pinjaman: Rp100.000.000
- Tenor: 36 bulan
- Suku bunga tahunan: 9% (0.75% per bulan)
- Formula: =PMT(0.09/12, 36, -100000000)
- Hasil: Rp3.140.656 (angsuran bulanan)

Interpretasi: Setiap bulan PT Maju Jaya bayar Rp3.140.656 selama 36 bulan untuk melunasi pembiayaan.
```

**Catatan Penting**:
- Tanda negatif pada `pv` sangat penting agar hasil angsuran bernilai positif (Excel menggunakan cash flow direction).
- Pastikan `rate` dan `nper` menggunakan periode yang sama (jika bulanan, rate harus rate per bulan, nper dalam bulan).

---

#### 1.2 PMT() Multi-Skenario (Analisis Dampak)
**Contoh Formula Perbandingan**:
```
Angsuran dengan tenor 36 bulan:
=PMT(0.09/12, 36, -100000000)

Angsuran dengan tenor 24 bulan:
=PMT(0.09/12, 24, -100000000)

Angsuran dengan tenor 48 bulan:
=PMT(0.09/12, 48, -100000000)
```

**Interpretasi**:
- Tenor lebih panjang → angsuran lebih kecil tapi total bunga lebih besar.
- Tenor lebih pendek → angsuran lebih besar tapi total bunga lebih kecil.

---

### Blok 2: Fungsi PV() - Present Value (Nilai Sekarang)

#### 2.1 PV() Dasar
**Definisi**: Menghitung nilai sekarang (jumlah uang hari ini) dari serangkaian pembayaran di masa depan dengan suku bunga tertentu.

**Sintaks**:
```
=PV(rate, nper, pmt, [fv], [type])
```

**Parameter**:
- rate: Suku bunga per periode
- nper: Jumlah periode pembayaran
- pmt: Besarnya pembayaran per periode. **Gunakan tanda negatif** untuk cash outflow.
- fv (opsional): Nilai masa depan yang ingin dicapai/tersisa
- type (opsional): 0 atau 1 (timing pembayaran)

**Return Value**: Nilai sekarang (uang pada hari ini yang setara dengan rangkaian pembayaran)

**Kegunaan Bisnis**:
- Validasi: Cek apakah angsuran yang dihitung benar (bandingkan dengan pinjaman awal)
- Evaluasi investasi: Berapa nilai investasi masa depan dalam nilai uang hari ini?
- Lease vs Buy: Membandingkan biaya operating lease dengan cash down payment

**Contoh**:
```
PT Maju Jaya - Validasi Angsuran
- Angsuran yang dihitung: Rp3.140.656/bulan
- Tenor: 36 bulan
- Suku bunga: 0.75% per bulan
- Formula: =PV(0.09/12, 36, -3140656)
- Hasil: ~Rp100.000.000

Interpretasi: Validasi berhasil! Angsuran Rp3.140.656 x 36 bulan = pinjaman Rp100 juta pada suku bunga 0.75% per bulan.
```

---

### Blok 3: Fungsi NPER() dan RATE()

#### 3.1 NPER() - Jumlah Periode
**Definisi**: Menghitung berapa lama (jumlah periode) diperlukan untuk melunasi pinjaman atau mencapai target dengan angsuran dan suku bunga tertentu.

**Sintaks**:
```
=NPER(rate, pmt, pv, [fv], [type])
```

**Kegunaan Bisnis**:
- Perencanaan: Berapa bulan/tahun untuk lunas jika angsuran ditingkatkan?
- Evaluasi: Apakah target pelunasan realistis dengan kapasitas pembayaran saat ini?

**Contoh**:
```
PT Maju Jaya - Perencanaan Pelunasan
- Pinjaman: Rp100.000.000
- Angsuran yang bisa dicicil: Rp4.000.000/bulan
- Suku bunga: 0.75% per bulan
- Formula: =NPER(0.09/12, -4000000, 100000000)
- Hasil: ~26.4 bulan (kurang dari 3 tahun)

Interpretasi: Dengan menaikkan angsuran ke Rp4 juta/bulan, pelunasan bisa dipercepat ke 26.4 bulan (hemat bunga).
```

#### 3.2 RATE() - Suku Bunga Per Periode
**Definisi**: Menghitung suku bunga per periode (iterative calculation) jika pinjaman, angsuran, dan tenor diketahui.

**Sintaks**:
```
=RATE(nper, pmt, pv, [fv], [type], [guess])
```

**Kegunaan Bisnis**:
- Analisis proposal: Berapa effective rate dari penawaran pembiayaan tertentu?
- Benchmarking: Bandingkan dengan rate pasar untuk negosiasi.

**Contoh**:
```
PT Maju Jaya - Reverse-Engineer Suku Bunga
- Pinjaman: Rp100.000.000
- Angsuran: Rp3.140.656/bulan
- Tenor: 36 bulan
- Formula: =RATE(36, -3140656, 100000000)
- Hasil: 0.0075 = 0.75% per bulan = 9% per tahun

Interpretasi: Penawaran pembiayaan ini menawarkan rate 9% per tahun (kompetitif di pasar).
```

---

### Blok 4: Fungsi IPMT() dan PPMT() - Komposisi Pembayaran

#### 4.1 IPMT() - Interest Payment (Bunga Per Periode)
**Definisi**: Menghitung besarnya bunga yang dibayar pada periode tertentu.

**Sintaks**:
```
=IPMT(rate, per, nper, pv, [fv], [type])
```

**Parameter**:
- per: Nomor periode (1=periode pertama, 2=periode kedua, dst)

**Return Value**: Besarnya bunga pada periode `per`

**Kegunaan Bisnis**:
- Accounting: Catat beban bunga per periode untuk laporan laba-rugi
- Tracking: Monitor bunga yang sudah dibayar vs akan dibayar
- Analisis: Lihat berapa persen pembayaran untuk bunga vs pokok

**Contoh**:
```
Bunga periode pertama: =IPMT(0.09/12, 1, 36, 100000000)
Bunga periode terakhir (ke-36): =IPMT(0.09/12, 36, 36, 100000000)

Hasil: Periode 1 bunga besar (Rp750.000), periode 36 bunga kecil (Rp19.500)
Interpretasi: Semakin lama, porsi bunga semakin kecil, porsi pokok semakin besar.
```

#### 4.2 PPMT() - Principal Payment (Pokok Per Periode)
**Definisi**: Menghitung besarnya pokok (principal) yang dibayar pada periode tertentu.

**Sintaks**:
```
=PPMT(rate, per, nper, pv, [fv], [type])
```

**Kegunaan Bisnis**:
- Track equity build: Lihat porsi pembayaran yang mengurangi saldo pinjaman (equity)
- Loan balance: Saldo pinjaman = saldo awal - kumulatif PPMT

**Contoh**:
```
Pokok periode pertama: =PPMT(0.09/12, 1, 36, 100000000)
Pokok periode terakhir: =PPMT(0.09/12, 36, 36, 100000000)

Hasil: Periode 1 pokok kecil (Rp2.390.656), periode 36 pokok besar (Rp3.121.156)
Interpretasi: Awal tenor mayoritas bayar bunga, akhir tenor mayoritas bayar pokok.
```

---

### Blok 5: Tabel Amortisasi (Aplikasi Terintegrasi)

#### 5.1 Struktur Tabel Amortisasi
Tabel amortisasi adalah riwayat lengkap pembayaran dari awal sampai akhir tenor, yang menunjukkan:
- Saldo pinjaman awal periode
- Angsuran (tetap setiap periode)
- Komposisi: berapa yang untuk bunga, berapa untuk pokok
- Saldo pinjaman akhir periode

**Kegunaan Bisnis**:
- Dokumentasi formal untuk kreditur/debitur
- Accounting: Source untuk pencatatan jurnal bunga dan pembayaran
- Planning: Lihat kapan saldo mencapai 50%, 75%, dst
- Negotiation: Jika ada perubahan suku bunga, hitung ulang amortisasi

#### 5.2 Template Tabel Amortisasi (Sheet PROSES)
Buat header baris 1:
- A: `Periode` (1, 2, 3, ..., n)
- B: `Saldo Awal` (sisa pinjaman awal periode)
- C: `Angsuran` (tetap, hasil PMT)
- D: `Bunga` (hasil IPMT)
- E: `Pokok` (hasil PPMT)
- F: `Saldo Akhir` (B - E)

**Rumus Baris 2 (Periode 1) - Asumsi Parameter di H1:H3**:
- H1: rate per bulan (0.09/12)
- H2: nper total (36)
- H3: pv pinjaman (-100000000)

```
A2: 1
B2: =ABS($H$3)          [Nilai mutlak PV untuk display positif]
C2: =PMT($H$1,$H$2,$H$3)   [Angsuran tetap]
D2: =IPMT($H$1,A2,$H$2,$H$3)  [Bunga periode 1]
E2: =PPMT($H$1,A2,$H$2,$H$3)  [Pokok periode 1]
F2: =B2-E2              [Saldo akhir = saldo awal - pokok dibayar]
```

**Rumus Baris 3+ (Periode 2 dan seterusnya)**:
```
A3: =A2+1               [Periode bertambah]
B3: =F2                 [Saldo awal = saldo akhir periode sebelumnya]
C3: =C2                 [Angsuran tetap, copy dari C2]
D3: =IPMT($H$1,A3,$H$2,$H$3)  [Bunga dihitung per periode ke-3]
E3: =PPMT($H$1,A3,$H$2,$H$3)  [Pokok dihitung per periode ke-3]
F3: =B3-E3              [Saldo akhir]
```

Copy rumus A3:F3 ke bawah sampai A38:F38 (36 periode).

**Validasi Tabel**:
- F38 (saldo akhir terakhir) harus = 0 (pinjaman terbayar penuh)
- SUM(D2:D37) = total bunga dibayar
- SUM(E2:E37) = total pokok dibayar = H3 (pinjaman awal)

---

## 📊 STRUKTUR DATA PRAKTIK

### Tabel Data Praktik: CV Komputer Indonesia - Loan Financing & Amortization

**Konteks Bisnis**:
CV Komputer Indonesia sedang merencanakan ekspansi dengan membeli perangkat kantor seharga Rp100 juta. Manajemen sedang mengevaluasi berbagai skenario pembiayaan (tenor 24/36/48 bulan, suku bunga berbeda). Anda berperan sebagai analis finansial untuk membuat proyeksi cash flow dan rekomendasi opsi pembiayaan terbaik.

### Struktur Input Data

| Kolom | Tipe | Isi | Contoh |
|-------|------|-----|---------|
| Loan Amount | Currency | Jumlah pinjaman | 100,000,000 |
| Annual Rate | Percentage | Suku bunga per tahun | 9% |
| Tenor (Months) | Integer | Durasi cicilan (bulan) | 36 |

### Sample Data Input (Sheet DATA_INPUT)

```
Skenario Pembiayaan CV Komputer Indonesia

Loan Details:
Loan Amount: 100000000
Annual Interest Rate: 0.09 (9%)
Tenor (Months): 36

Calculated Parameters:
Monthly Rate: =Loan Amount / 12 = 0.0075
```

### Output KPI Wajib (Sheet DASHBOARD)

| KPI | Formula | Contoh Hasil |
|-----|---------|-------------|
| Angsuran Bulanan | `=PMT(rate, nper, -pv)` | Rp 3,140,656 |
| Total Pembayaran | `=PMT(...) * nper` | Rp 113,063,616 |
| Total Bunga | `=Total Pembayaran - Loan Amount` | Rp 13,063,616 |
| Effective Cost % | `=(Total Bunga / Loan Amount) * 100` | 13.06% |

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Penerapan Dasar Fungsi Finansial
*Fokus: PMT, PV, NPER, RATE - Durasi: 40 menit*

#### Latihan 1.1: Hitung Angsuran dengan PMT()
**Waktu**: 10 menit

**Instruksi**:
1. Buat sheet `DATA_INPUT` berisi:
   - Loan Amount: 100,000,000
   - Annual Rate: 0.09 (9%)
   - Tenor: 36 bulan
2. Di sheet `PROSES`, hitung monthly rate di sel H1: `=0.09/12`
3. Di H2, hitung monthly payment: `=PMT(H1, 36, -100000000)`

**Pertanyaan Analisis**:
- Berapa angsuran bulanan yang harus dibayar?
- Apakah angsuran ini feasible untuk cash flow perusahaan?

---

#### Latihan 1.2: Validasi dengan PV()
**Waktu**: 8 menit

**Instruksi**:
1. Di H3, hitung PV untuk validasi: `=PV(H1, 36, -H2)`
2. Bandingkan hasil H3 dengan pinjaman awal (H1).

**Pertanyaan Analisis**:
- Apakah PV sama dengan pinjaman awal? (Harus = Rp100 juta)
- Apa arti validasi ini?

---

#### Latihan 1.3: Analisis Multi-Tenor dengan NPER()
**Waktu**: 12 menit

**Instruksi**:
1. Di H5, hitung tenor jika angsuran diturunkan jadi Rp3.5 juta: `=NPER(H1, -3500000, 100000000)`
2. Di H6, hitung tenor jika angsuran dinaikkan jadi Rp3.5 juta: `=NPER(H1, -4000000, 100000000)`

**Pertanyaan Analisis**:
- Jika angsuran diturunkan, berapa lama tenor pinjaman?
- Jika angsuran dinaikkan, berapa lama tenor pinjaman?
- Apa trade-off antara tenor panjang vs pendek?

---

### LEVEL 2: Tabel Amortisasi Lengkap & Analisis
*Fokus: IPMT, PPMT, Amortization Schedule - Durasi: 50 menit*

#### Latihan 2.1: Build Amortization Table
**Waktu**: 20 menit

**Instruksi**:
1. Di sheet `PROSES`, buat header amortisasi (baris 1):
   - A: Periode
   - B: Saldo Awal
   - C: Angsuran
   - D: Bunga
   - E: Pokok
   - F: Saldo Akhir

2. Isi baris 2 (periode 1):
   - A2: 1
   - B2: `=100000000`
   - C2: `=PMT($H$1, 36, -100000000)`
   - D2: `=IPMT($H$1, A2, 36, -100000000)`
   - E2: `=PPMT($H$1, A2, 36, -100000000)`
   - F2: `=B2 - E2`

3. Copy rumus A2:F2 ke A3:A37 (36 periode total).
   - Modifikasi A3: `=A2+1`
   - Modifikasi B3: `=F2`

**Validasi**:
- Cek F37 (saldo akhir): harus = 0 atau mendekati 0 (Rp < 1,000)
- Total bunga: `=SUM(D2:D37)`

---

#### Latihan 2.2: Analisis Komposisi Pembayaran
**Waktu**: 15 menit

**Instruksi**:
1. Di DASHBOARD, hitung:
   - Total Pembayaran: `=SUM(C2:C37)`
   - Total Bunga: `=SUM(D2:D37)`
   - Total Pokok: `=SUM(E2:E37)`
   - Verifikasi: Total Pokok harus = Loan Amount

2. Buat insight teks:
   - Berapa % pembayaran untuk bunga? (=Total Bunga / Total Pembayaran * 100)
   - Berapa % pembayaran untuk pokok?

**Pertanyaan Analisis**:
- Di periode awal, pembayaran mayoritas untuk bunga atau pokok?
- Di periode akhir, pembayaran mayoritas untuk bunga atau pokok?
- Apa implikasi untuk accounting (kapan expense terbesar)?

---

#### Latihan 2.3: Skenario Perbandingan & Rekomendasi
**Waktu**: 15 menit

**Instruksi**:
1. Buat 3 skenario di sheet terpisah (atau tab) untuk tenor berbeda:
   - Skenario A: 24 bulan @ 9% pa
   - Skenario B: 36 bulan @ 9% pa (baseline)
   - Skenario C: 48 bulan @ 9% pa

2. Untuk masing-masing, hitung:
   - Angsuran bulanan
   - Total bunga
   - Effective cost %

3. Buat tabel perbandingan di DASHBOARD:

   | Skenario | Tenor | Angsuran Bulanan | Total Bunga | Cost % |
   |----------|-------|-----------------|-----------|--------|
   | A | 24 | ? | ? | ? |
   | B | 36 | ? | ? | ? |
   | C | 48 | ? | ? | ? |

**Pertanyaan Analisis & Rekomendasi**:
- Skenario mana yang paling cost-efficient (total bunga terendah)?
- Skenario mana yang paling cash-friendly (angsuran terendah)?
- Jika cash flow CV Komputer terbatas, skenario mana yang direkomendasikan?
- Apa trade-off antara cost efficiency dan cash flow flexibility?

**Insight Wajib**:
- Tulis minimal 3 insight berbasis data
- Tulis minimal 2 rekomendasi untuk management

---

## 💡 Tips & Best Practices

### Prinsip Umum Fungsi Finansial
1. **Tanda Cash Flow Sangat Penting**
   - Excel menggunakan konvensi: inflow = positif, outflow = negatif
   - Untuk PMT: pv sebaiknya negatif agar PMT positif (angsuran = outflow)
   - Untuk PV: pmt sebaiknya negatif agar PV positif (pembayaran = outflow)

2. **Konsistensi Rate & NPER**
   - Jika rate = 9% per tahun, gunakan rate = 0.09/12 untuk periode bulanan
   - Jika rate = 9% per tahun, gunakan rate = 0.09 dan nper dalam tahun
   - **Jangan mix**: rate bulanan dengan nper tahunan

3. **Pembulatan untuk Rupiah**
   - Gunakan `=ROUND(PMT(...), 0)` untuk angsuran tanpa desimal
   - Gunakan `=ROUND(IPMT(...), 0)` untuk bunga per periode
   - Desimal akan mengecil di periode akhir (saldo akhir ≈ 0 tapi bukan 0 persis)

4. **Validasi Tabel Amortisasi**
   - Cek: F (saldo akhir terakhir) ≈ 0 (toleransi Rp1 untuk rounding)
   - Cek: SUM(E2:En) = Loan Amount (pokok)
   - Cek: PMT bulanan tetap sama dari baris ke baris

### Kesalahan Umum
- Lupa tanda negatif pada pv/pmt → hasil berlawanan
- Mencampur periode rate (monthly rate tapi nper tahunan)
- Tidak memvalidasi tabel amortisasi → cascade error hingga baris terakhir

---

## 📚 Referensi & Checklist Penilaian

### Checklist Implementasi
- [ ] Sheet DATA_INPUT berisi parameter pinjaman (amount, rate, tenor)
- [ ] Sheet PROSES berisi amortisasi 36 periode dengan formula PMT, IPMT, PPMT
- [ ] Saldo akhir terakhir (F37) ≈ 0
- [ ] Sheet DASHBOARD berisi KPI minimal: angsuran, total pembayaran, total bunga
- [ ] Minimal 3 insight berbasis data
- [ ] Minimal 2 rekomendasi untuk management

### Kompetensi yang Terasah (Learning Outcomes Tercapai)
- ✅ **C3 (Apply)**: Mahasiswa bisa mengaplikasikan PMT, PV, NPER, RATE, IPMT, PPMT pada kasus nyata
- ✅ **C4 (Analyze)**: Mahasiswa bisa menganalisis dampak rate/tenor pada total cost dan cash flow
