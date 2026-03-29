# MID PROJECT: Integrasi Excel Dasar (Week 1 - Week 7)

## 📋 METADATA PROYEK

### Informasi Dasar
- **Nama Kegiatan**: Mid Project Praktikum Excel Dasar
- **Cakupan Materi**: Week 1 sampai Week 7
- **Durasi Pengerjaan**: 1 pertemuan (120 menit) + finalisasi mandiri
- **Level**: Beginner
- **Konteks**: Bisnis/Ekonomi (CV Komputer Indonesia)

### Tujuan Mid Project
Mid project ini digunakan untuk **menilai dan mengevaluasi pemahaman mahasiswa** selama 7 pertemuan, terutama kemampuan:
1. Menyusun worksheet rapi dan terstruktur (Week 1)
2. Menggunakan formula dasar perhitungan (Week 2)
3. Mengolah teks dan agregasi numerik (Week 3)
4. Mengolah fungsi tanggal dan waktu (Week 4)
5. Menerapkan logika keputusan IF (Week 5)
6. Menganalisis data berbasis kriteria (Week 6)
7. Menghubungkan transaksi dengan master data via lookup (Week 7)

### Cerita Besar Project (Supaya Tidak Bingung)
CV Komputer Indonesia sedang menyiapkan laporan evaluasi tengah semester untuk manajemen. Data yang dimiliki masih terpisah: ada data mentah transaksi/proyek dan ada data master produk. Anda berperan sebagai analis data pemula yang harus mengubah data mentah menjadi laporan siap keputusan.

Hasil akhir yang diharapkan:
1. Data mentah berhasil diolah tanpa error formula.
2. Dashboard KPI menampilkan kondisi bisnis secara cepat.
3. Ada insight dan rekomendasi operasional berbasis angka.

---

## 🧭 ATURAN PEMBAGIAN KASUS

### Penentuan Soal
- **Kasus A**: untuk mahasiswa dengan **ID Genap**
- **Kasus B**: untuk mahasiswa dengan **ID Ganjil**

### Ketentuan Umum
- Dilarang menukar kasus antar mahasiswa.
- Semua formula harus terlihat (tidak di-hide).
- Gunakan nama file sesuai format:
  - **NAMA_NIM_MidProject_CaseA.xlsx**
  - **NAMA_NIM_MidProject_CaseB.xlsx**

### Struktur File Wajib (Agar Tidak Bingung)
Gunakan **3 sheet** berikut untuk kedua kasus:
1. **DATA_INPUT**: hanya berisi data mentah (copy data soal).
2. **PROSES**: semua perhitungan formula dilakukan di sini.
3. **DASHBOARD**: ringkasan KPI, status, insight, dan rekomendasi.

### Penjelasan Sangat Jelas: Sheet PROSES Itu Isinya Apa?
Sheet **PROSES** adalah tempat kerja utama. Di sheet ini Anda **tidak input data mentah baru**, tetapi:
1. Menyalin kolom sumber dari DATA_INPUT.
2. Menambahkan kolom hasil lookup (nama produk, harga, kategori).
3. Menambahkan kolom hasil hitung (total sales/biaya, working days, umur transaksi, dst).
4. Menambahkan kolom validasi/status (OK/Perlu Koreksi, High/Normal).
5. Menambahkan kolom bantu teks/tanggal bila diminta (LEFT/RIGHT/MID, DAY/MONTH/YEAR).

Ringkasnya:
- DATA_INPUT = data mentah
- PROSES = data olahan + semua formula
- DASHBOARD = ringkasan KPI + insight

Urutan kerja yang wajib diikuti:
1. Input data ke DATA_INPUT.
2. Susun header kolom di PROSES.
3. Isi formula di baris pertama data.
4. Copy formula ke seluruh baris.
5. Buat KPI di DASHBOARD.
6. Tulis insight dan rekomendasi berbasis angka.

---

## 🔗 KESINAMBUNGAN DATA (WEEK 1 - 7)

Agar konsisten dengan praktikum sebelumnya, data mid project ini mengambil pola dari materi yang sudah dipakai:
- **Week 3**: Invoice ID `INV-0326-001` s.d. `INV-0326-008`, nama sales (Budi, Ani, Rudi)
- **Week 4**: Data karyawan `EMP-001` s.d. `EMP-008` dan tanggal project
- **Week 5**: Data performa sales, status `Lunas/Hutang`, logika eligibility
- **Week 6**: Segmentasi data berbasis `SUMIF/COUNTIF/AVERAGEIF`
- **Week 7**: Master produk `PRD-001` s.d. `PRD-008`, lookup VLOOKUP/HLOOKUP

---

## ✅ KOMPETENSI YANG WAJIB MUNCUL DI MID PROJECT

Minimal formula/fungsi yang harus digunakan:
- **Formula dasar**: `+`, `-`, `*`, `/`
- **Agregasi**: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`
- **Teks**: `LEFT`, `RIGHT`, `MID`, `CONCATENATE` / `&`
- **Tanggal**: `TODAY`, `YEAR`, `MONTH`, `DAY`, `DATEDIF`, `NETWORKDAYS`
- **Logika**: `IF`, `AND`, `OR`
- **Kriteria**: `SUMIF`, `COUNTIF`, `AVERAGEIF`
- **Lookup**: `VLOOKUP`, `HLOOKUP`

Catatan:
- **INDEX-MATCH** boleh digunakan sebagai nilai tambah (opsional), tetapi tidak wajib.

---

## 🅰️ KASUS A (ID GENAP)

### Judul Kasus
**Analisis Penjualan Maret 2026 dan Validasi Master Produk**

### Latar Belakang
Manajemen CV Komputer Indonesia ingin mengevaluasi performa penjualan dan kualitas data transaksi. Data transaksi harian harus dihubungkan dengan master produk untuk menghasilkan laporan operasional yang akurat.

### Narasi Kasus A (Apa yang Dikerjakan dan Dicapai)
Anda diminta membangun alur laporan penjualan dari awal sampai siap presentasi manajemen:
1. Melengkapi data transaksi dengan master produk.
2. Menandai transaksi yang bermasalah (kode tidak valid).
3. Menghitung KPI sales dan capaian target bulanan.
4. Menuliskan insight dan rekomendasi yang bisa langsung ditindak.

### Data yang Digunakan

#### A. Data Transaksi (Input)

| Invoice ID | Sales Person | Kode Produk | Region | Channel | Tanggal Transaksi | Qty |
|---|---|---|---|---|---|---|
| INV-0326-001 | Budi Santoso | PRD-001 | Malang | Online | 2026-03-01 | 2 |
| INV-0326-002 | Ani Wijaya | PRD-004 | Surabaya | Offline | 2026-03-03 | 4 |
| INV-0326-003 | Rudi Firmansah | PRD-003 | Malang | Online | 2026-03-05 | 3 |
| INV-0326-004 | Budi Santoso | PRD-002 | Blitar | Online | 2026-03-08 | 1 |
| INV-0326-005 | Ani Wijaya | PRD-008 | Kediri | Offline | 2026-03-10 | 6 |
| INV-0326-006 | Rudi Firmansah | PRD-006 | Malang | Online | 2026-03-15 | 2 |
| INV-0326-007 | Budi Santoso | PRD-005 | Surabaya | Online | 2026-03-20 | 5 |
| INV-0326-008 | Ani Wijaya | PRD-999 | Malang | Offline | 2026-03-25 | 2 |

#### B. Master Produk (Referensi)

| Kode Produk | Nama Produk | Harga Unit | Kategori |
|---|---|---:|---|
| PRD-001 | Laptop Office 14 | 8500000 | Laptop |
| PRD-002 | Printer Inkjet X | 2200000 | Printer |
| PRD-003 | Monitor 24 FHD | 3200000 | Monitor |
| PRD-004 | Laptop Pro 15 | 12500000 | Laptop |
| PRD-005 | Keyboard Mechanical | 950000 | Aksesoris |
| PRD-006 | Printer Laser Pro | 4800000 | Printer |
| PRD-007 | Mouse Wireless | 250000 | Aksesoris |
| PRD-008 | UPS 1200VA | 1750000 | Power |

#### C. Target Bulanan (Horizontal Table untuk HLOOKUP)

| Bulan | Jan | Feb | Mar | Apr | Mei | Jun |
|---|---:|---:|---:|---:|---:|---:|
| Target Sales | 70000000 | 75000000 | 90000000 | 85000000 | 95000000 | 100000000 |
| Target Margin (%) | 18 | 18 | 20 | 20 | 21 | 22 |

### Tugas Pengerjaan

### Panduan Eksekusi Detail (Wajib Ikuti Urutan)

#### Step A1 - Siapkan Sheet PROSES
Di sheet PROSES, buat header kolom:
- A: Invoice ID
- B: Sales Person
- C: Kode Produk
- D: Region
- E: Channel
- F: Tanggal Transaksi
- G: Qty
- H: Nama Produk
- I: Harga Unit
- J: Kategori
- K: Total Sales
- L: Validasi
- M: Prefix Invoice
- N: Nomor Urut Invoice
- O: Hari
- P: Bulan
- Q: Tahun
- R: Umur Transaksi (Hari)

Copy data A:G dari DATA_INPUT ke PROSES.

#### Step A2 - Formula Lookup (Ketik di baris 2)
Di H2:
```excel
=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,2,FALSE),"Kode Tidak Valid")
```
Di I2:
```excel
=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,3,FALSE),0)
```
Di J2:
```excel
=IFERROR(VLOOKUP(C2,DATA_INPUT!$M$2:$P$9,4,FALSE),"Unknown")
```
Setelah 3 formula benar, copy ke bawah sampai baris data terakhir.

#### Step A3 - Formula Perhitungan dan Teks
Di K2:
```excel
=G2*I2
```
Di L2:
```excel
=IF(H2="Kode Tidak Valid","Perlu Koreksi","OK")
```
Di M2:
```excel
=LEFT(A2,3)
```
Di N2:
```excel
=RIGHT(A2,3)
```

#### Step A4 - Formula Tanggal
Di O2:
```excel
=DAY(F2)
```
Di P2:
```excel
=MONTH(F2)
```
Di Q2:
```excel
=YEAR(F2)
```
Di R2:
```excel
=DATEDIF(F2,TODAY(),"D")
```

#### Step A5 - Dashboard KPI Minimal
Wajib ada di DASHBOARD:
- Total Sales: `=SUM(PROSES!K2:K9)`
- Rata-rata Sales: `=AVERAGE(PROSES!K2:K9)`
- Min Sales: `=MIN(PROSES!K2:K9)`
- Max Sales: `=MAX(PROSES!K2:K9)`
- Jumlah Transaksi: `=COUNT(PROSES!K2:K9)`
- Valid: `=COUNTIF(PROSES!L2:L9,"OK")`
- Invalid: `=COUNTIF(PROSES!L2:L9,"Perlu Koreksi")`
- Sales per kategori (contoh Laptop): `=SUMIF(PROSES!J2:J9,"Laptop",PROSES!K2:K9)`
- Transaksi Online: `=COUNTIF(PROSES!E2:E9,"Online")`
- Avg Sales Malang: `=AVERAGEIF(PROSES!D2:D9,"Malang",PROSES!K2:K9)`

Target Maret dengan HLOOKUP:
```excel
=HLOOKUP("Mar",DATA_INPUT!$T$1:$Z$3,2,FALSE)
```

Status pencapaian:
```excel
=IF([sel_total_sales]>=[sel_target_maret],"ON TRACK","BELOW TARGET")
```

#### Step A6 - Insight Wajib
Tulis minimal:
- 3 insight (harus menyebut angka KPI)
- 3 rekomendasi operasional (harus relevan dengan insight)

#### LEVEL 1 - Build Data (40 menit)
1. Lengkapi kolom `Nama Produk`, `Harga Unit`, `Kategori` dengan `VLOOKUP` + `IFERROR`.
2. Hitung `Total Sales = Qty * Harga Unit`.
3. Buat kolom `Validasi`:
   - `OK` jika kode produk valid
   - `Perlu Koreksi` jika kode tidak ditemukan.
4. Buat kolom `Kode Singkat` dari Invoice:
   - `LEFT` untuk prefix `INV`
   - `RIGHT` untuk nomor urut transaksi.

#### LEVEL 2 - Analysis Dashboard (50 menit)
1. Buat KPI utama:
   - Total Sales (`SUM`)
   - Rata-rata transaksi (`AVERAGE`)
   - Nilai transaksi minimum/maksimum (`MIN`/`MAX`)
   - Jumlah transaksi valid/invalid (`COUNTIF`)
2. Buat analisis berbasis kriteria:
   - Total sales per kategori (`SUMIF`)
   - Jumlah transaksi per channel (`COUNTIF`)
   - Rata-rata sales region Malang (`AVERAGEIF`)
3. Buat analisis tanggal:
   - Hari transaksi (`DAY`)
   - Bulan transaksi (`MONTH`)
   - Umur transaksi terhadap hari ini (`DATEDIF`/`TODAY`)
4. Ambil target bulan Maret dengan `HLOOKUP`, lalu buat status:
   - `ON TRACK` jika total sales >= target
   - `BELOW TARGET` jika belum tercapai (`IF`).
5. Tulis minimal 3 insight bisnis dan 3 rekomendasi.

---

## 🅱️ KASUS B (ID GANJIL)

### Judul Kasus
**Analisis Produktivitas Karyawan, Timeline Proyek, dan Kebutuhan Pengadaan**

### Latar Belakang
Divisi HR dan Operasional CV Komputer Indonesia membutuhkan laporan gabungan untuk mengevaluasi produktivitas karyawan, durasi proyek, serta estimasi kebutuhan pengadaan perangkat berdasarkan aktivitas proyek.

### Narasi Kasus B (Apa yang Dikerjakan dan Dicapai)
Anda diminta membuat laporan gabungan HR-operasional:
1. Mengukur masa kerja dan durasi kerja proyek tiap karyawan.
2. Menghubungkan kebutuhan produk ke master data untuk menghitung biaya.
3. Menentukan prioritas tindak lanjut berdasarkan kondisi data dan timeline.
4. Menyajikan KPI dan rekomendasi agar manajemen tahu area prioritas.

### Data yang Digunakan

#### A. Data Karyawan & Proyek (Input)

| ID Karyawan | Nama Karyawan | Hire Date | Posisi | Project Start | Project End | Kode Produk Utama | Qty Kebutuhan |
|---|---|---|---|---|---|---|---:|
| EMP-001 | Budi Santoso | 2020-05-15 | Sales Manager | 2026-03-01 | 2026-03-31 | PRD-001 | 2 |
| EMP-002 | Ani Wijaya | 2019-08-20 | Finance Manager | 2026-03-01 | 2026-03-22 | PRD-004 | 1 |
| EMP-003 | Rudi Firmansah | 2022-01-10 | Tech Support | 2026-03-05 | 2026-03-31 | PRD-003 | 3 |
| EMP-004 | Siti Nurhaliza | 2021-03-01 | HR Coordinator | 2026-03-01 | 2026-03-15 | PRD-002 | 2 |
| EMP-005 | Ahmad Dahlan | 2023-06-12 | Junior Developer | 2026-03-08 | 2026-03-31 | PRD-006 | 2 |
| EMP-006 | Dewi Lestari | 2018-11-05 | Director | 2026-03-01 | 2026-03-31 | PRD-008 | 1 |
| EMP-007 | Hendra Wijaya | 2021-07-22 | Data Analyst | 2026-03-15 | 2026-03-31 | PRD-005 | 4 |
| EMP-008 | Linda Putri | 2022-02-14 | Customer Service | 2026-03-01 | 2026-03-25 | PRD-999 | 1 |

#### B. Master Produk (Referensi)

| Kode Produk | Nama Produk | Harga Unit | Kategori |
|---|---|---:|---|
| PRD-001 | Laptop Office 14 | 8500000 | Laptop |
| PRD-002 | Printer Inkjet X | 2200000 | Printer |
| PRD-003 | Monitor 24 FHD | 3200000 | Monitor |
| PRD-004 | Laptop Pro 15 | 12500000 | Laptop |
| PRD-005 | Keyboard Mechanical | 950000 | Aksesoris |
| PRD-006 | Printer Laser Pro | 4800000 | Printer |
| PRD-007 | Mouse Wireless | 250000 | Aksesoris |
| PRD-008 | UPS 1200VA | 1750000 | Power |

#### C. Target KPI Bulanan (Horizontal Table untuk HLOOKUP)

| KPI | Jan | Feb | Mar | Apr | Mei | Jun |
|---|---:|---:|---:|---:|---:|---:|
| Target Working Days Avg | 18 | 19 | 20 | 20 | 21 | 21 |
| Target Cost Efficiency (%) | 80 | 82 | 85 | 85 | 86 | 87 |

### Tugas Pengerjaan

### Panduan Eksekusi Detail (Wajib Ikuti Urutan)

#### Step B1 - Siapkan Sheet PROSES
Di sheet PROSES, buat header kolom:
- A: ID Karyawan
- B: Nama Karyawan
- C: Hire Date
- D: Posisi
- E: Project Start
- F: Project End
- G: Kode Produk Utama
- H: Qty Kebutuhan
- I: Tenure (Tahun)
- J: Hire Year
- K: Hire Month
- L: Hire Day
- M: Working Days
- N: Nama Produk
- O: Harga Unit
- P: Kategori
- Q: Total Kebutuhan Biaya
- R: Status Validasi Produk
- S: Prioritas

Copy data A:H dari DATA_INPUT ke PROSES.

#### Step B2 - Formula Tanggal dan Durasi
Di I2:
```excel
=DATEDIF(C2,TODAY(),"Y")
```
Di J2:
```excel
=YEAR(C2)
```
Di K2:
```excel
=MONTH(C2)
```
Di L2:
```excel
=DAY(C2)
```
Di M2:
```excel
=NETWORKDAYS(E2,F2)
```

#### Step B3 - Formula Lookup Produk
Di N2:
```excel
=IFERROR(VLOOKUP(G2,DATA_INPUT!$M$2:$P$9,2,FALSE),"Kode Tidak Valid")
```
Di O2:
```excel
=IFERROR(VLOOKUP(G2,DATA_INPUT!$M$2:$P$9,3,FALSE),0)
```
Di P2:
```excel
=IFERROR(VLOOKUP(G2,DATA_INPUT!$M$2:$P$9,4,FALSE),"Unknown")
```

#### Step B4 - Formula Biaya, Validasi, Prioritas
Di Q2:
```excel
=H2*O2
```
Di R2:
```excel
=IF(N2="Kode Tidak Valid","Perlu Koreksi","OK")
```
Di S2:
```excel
=IF(OR(M2>20,R2="Perlu Koreksi"),"High","Normal")
```

#### Step B5 - Dashboard KPI Minimal
Wajib ada di DASHBOARD:
- Avg Tenure: `=AVERAGE(PROSES!I2:I9)`
- Avg Working Days: `=AVERAGE(PROSES!M2:M9)`
- Total Biaya: `=SUM(PROSES!Q2:Q9)`
- Invalid Produk: `=COUNTIF(PROSES!R2:R9,"Perlu Koreksi")`
- High Priority: `=COUNTIF(PROSES!S2:S9,"High")`
- Biaya per kategori (contoh Laptop): `=SUMIF(PROSES!P2:P9,"Laptop",PROSES!Q2:Q9)`
- Jumlah posisi tertentu: `=COUNTIF(PROSES!D2:D9,"Tech Support")`
- Avg working days posisi tertentu: `=AVERAGEIF(PROSES!D2:D9,"Sales Manager",PROSES!M2:M9)`

Target KPI Maret:
```excel
=HLOOKUP("Mar",DATA_INPUT!$T$1:$Z$3,2,FALSE)
```

Status KPI:
```excel
=IF([sel_avg_working_days]<=[sel_target_working_days],"ON TRACK","BELOW TARGET")
```

#### Step B6 - Insight Wajib
Tulis minimal:
- 3 insight (harus berbasis angka KPI)
- 3 rekomendasi operasional

#### LEVEL 1 - Build Data (40 menit)
1. Hitung masa kerja:
   - `Tenure (tahun)` dengan `DATEDIF`.
   - Ekstrak `Hire Year`, `Hire Month`, `Hire Day`.
2. Hitung durasi proyek:
   - `Working Days` dengan `NETWORKDAYS`.
3. Lengkapi data produk dengan `VLOOKUP` + `IFERROR`:
   - `Nama Produk`, `Harga Unit`, `Kategori`.
4. Hitung `Total Kebutuhan Biaya = Qty Kebutuhan * Harga Unit`.
5. Buat `Status Validasi Produk` (`OK` / `Perlu Koreksi`).

#### LEVEL 2 - Analysis Dashboard (50 menit)
1. Buat KPI utama:
   - Rata-rata tenure (`AVERAGE`)
   - Rata-rata working days (`AVERAGE`)
   - Total biaya kebutuhan (`SUM`)
   - Jumlah data invalid produk (`COUNTIF`)
2. Buat analisis berbasis kriteria:
   - Total biaya per kategori (`SUMIF`)
   - Jumlah karyawan per posisi tertentu (`COUNTIF`)
   - Rata-rata working days untuk posisi tertentu (`AVERAGEIF`)
3. Buat logika prioritas dengan `IF` + `OR`:
   - `High` jika working days > 20 atau data produk invalid
   - `Normal` selain itu.
4. Ambil target Maret dengan `HLOOKUP`, lalu evaluasi status KPI (`ON TRACK` / `BELOW TARGET`).
5. Tulis minimal 3 insight bisnis dan 3 rekomendasi.

---

## 🧾 RUBRIK PENILAIAN MID PROJECT

| Komponen | Indikator | Bobot |
|---|---|---:|
| **Akurasi Formula (C3)** | Rumus benar, tanpa error, fungsi week1-week7 digunakan sesuai konteks | 50% |
| **Analisis & Insight (C4)** | Insight berbasis data, rekomendasi operasional jelas dan relevan | 30% |
| **Kerapian & Struktur File** | Layout rapi, label jelas, worksheet mudah dibaca, format konsisten | 20% |

**Total**: 100%

### Kriteria Kelulusan
- Nilai akhir minimal **70**
- Tidak ada error kritis formula (`#N/A` yang tidak ditangani, `#VALUE!`, `#REF!`)

---

## 📌 CHECKLIST WAJIB PENGUMPULAN

- [ ] Menggunakan fungsi dari week1-week7 (minimal 1 kali per kelompok fungsi utama)
- [ ] Dashboard ringkasan selesai (KPI + status + insight)
- [ ] Minimal 3 insight dan 3 rekomendasi
- [ ] Formula terlihat jelas (tidak hide)
- [ ] Penamaan file sesuai ketentuan

---

## 📖 CATATAN UNTUK PENGAJAR/ASISTEN

- Pastikan mahasiswa mengerjakan sesuai pembagian ID genap/ganjil.
- Fokus evaluasi bukan hanya hasil akhir angka, tetapi juga:
  - Struktur logika formula
  - Konsistensi referensi range
  - Kemampuan interpretasi bisnis
- Beri nilai tambah jika mahasiswa menggunakan pendekatan alternatif yang benar (misalnya INDEX-MATCH untuk eksplorasi mandiri).

---

## End of Document
