# Latihan Pertemuan 11: Data Management and Cleansing for Accountants
## 📋 METADATA DAN OBE FRAMEWORK

### Informasi Dasar
*   **Kode Modul**: EXCEL-011
*   **Pertemuan**: 11
*   **Durasi Praktik**: 90 menit
*   **Tingkat Kesulitan**: Dasar (Basic)
*   **Prasyarat**: Mahasiswa telah menguasai Week 1 s.d. Week 10

### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:
1.  **C3 (Apply)** - Menerapkan teknik manajemen data dalam konteks akuntansi:
    *   Melakukan **Import Data** dari file eksternal (.csv/.txt).
    *   Melakukan **Data Cleansing** menggunakan *Text to Columns* dan *Remove Duplicates*.
    *   Mengorganisir data menggunakan fitur **Format as Table** (materi yang ditarik dari Week 12).
2.  **C4 (Analyze)** - Menganalisis dan menyajikan informasi keuangan secara terstruktur:
    *   Menggunakan **Sorting** dan **Filtering** untuk menemukan transaksi spesifik.
    *   Menerapkan **Group and Ungroup** untuk membuat ringkasan laporan (outlining).

---

## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR

### Blok 1: Get & Clean Data (Persiapan Data Akuntansi)
**Definisi**: Proses mengambil data mentah dari sistem akuntansi lain dan membersihkannya agar siap diolah di Excel.
*   **Import External Data**: Membuka file berbasis teks (.txt atau .csv) ke dalam spreadsheet.
*   **Text to Columns**: Memecah satu sel berisi teks panjang (misal: "KODE-NAMA AKUN") menjadi kolom yang terpisah.
*   **Remove Duplicates**: Memastikan integritas data dengan menghapus baris transaksi yang terinput ganda secara tidak sengaja.

### Blok 2: Organize & Analyze (Strukturasi Data)
**Definisi**: Mengubah data mentah menjadi informasi yang mudah dicari dan diurutkan.
*   **Format as Table**: Mengaktifkan fitur tabel otomatis untuk mempermudah manajemen referensi data.
*   **Sorting & Filtering**: Mengurutkan data berdasarkan nominal/tanggal dan menyaring akun tertentu (misal: hanya melihat "Beban Gaji").

### Blok 3: Data Outlining (Presentasi Laporan)
**Definisi**: Meringkas data transaksi yang detail menjadi tampilan per kategori menggunakan fitur **Group/Ungroup**.

---

## 📊 STRUKTUR DATA PRAKTIK
### Kasus Bisnis: Mutasi Rekening Bank CV XYZ
**Konteks**: Anda menerima data mutasi bank dalam format teks (.csv) yang berantakan dari sistem perbankan. Data ini berisi transaksi ganda dan informasi yang tergabung dalam satu kolom. Anda diminta untuk merapikannya menjadi laporan yang siap dipresentasikan kepada Manajer Keuangan.

---

## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)

### LEVEL 1: Data Cleansing & Import
*Fokus: Import, Text to Columns, Remove Duplicates - Durasi: 40 menit*

**Instruksi**:
1.  Salin data mentah yang disediakan ke Excel.
2.  Gunakan **Text to Columns** dengan *Delimited* (Koma) untuk memisahkan Tanggal, Keterangan, dan Nominal.
3.  Gunakan **Remove Duplicates** untuk menghapus transaksi yang terekam dua kali.

**Pertanyaan Analisis**:
*   Mengapa proses pembersihan data krusial sebelum membuat jurnal akuntansi?

---

### LEVEL 2: Data Organizing & Reporting
*Fokus: Format as Table, Sort, Filter, Grouping - Durasi: 50 menit*

**Instruksi**:
1.  Ubah data yang sudah bersih menjadi **Format as Table**.
2.  Lakukan **Sorting** berdasarkan nominal terbesar ke terkecil untuk mendeteksi pengeluaran signifikan.
3.  Gunakan **Filter** untuk hanya menampilkan transaksi "Beban Operasional".
4.  Gunakan fitur **Group** pada baris transaksi mingguan untuk menyembunyikan detail harian (Outlining).

**Validasi**:
*   Pastikan total saldo akhir sesuai dengan data bank asli.
*   Checklist: Apakah tabel sudah memiliki filter otomatis di setiap kolom?

---

## 💡 Tips & Best Practices
1.  **Gunakan Delimiter yang Tepat**: Selalu periksa apakah data menggunakan koma (,) atau titik koma (;) saat melakukan *Text to Columns*.
2.  **Backup Data Asli**: Selalu simpan satu salinan data mentah sebelum melakukan *Remove Duplicates*.
3.  **Tanda Tabel**: Memberikan nama pada *Table* di Excel (Table Name) akan sangat memudahkan saat bekerja dengan rumus di pertemuan selanjutnya.

---

## 📚 Checklist Penilaian
*   [ ] Berhasil melakukan Import/Copy data teks ke kolom Excel.
*   [ ] Data bersih dari duplikasi.
*   [ ] Menggunakan *Format as Table* dengan desain profesional.
*   [ ] Mampu mendemonstrasikan hasil filter transaksi spesifik.
*   [ ] Menggunakan *Group/Ungroup* untuk meringkas tampilan laporan.