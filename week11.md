# Latihan Pertemuan 11: Visualisasi Tren dan Audit Data Keuangan
## 📋 METADATA DAN OBE FRAMEWORK
### Informasi Dasar
Kode Modul: EXCEL-011
Pertemuan: 11
Durasi Praktik: 90 menit
Tingkat Kesulitan: Dasar - Menengah (Basic to Intermediate)
Prasyarat: Mahasiswa telah menguasai Week 1 s.d. Week 10 (khususnya Format as Table)
.
### Learning Outcomes (Capaian Pembelajaran) - OBE Approach
Setelah menyelesaikan latihan ini, mahasiswa mampu:
C3 (Apply) - Menerapkan teknik visualisasi data mikro dan kontrol input:
Menggunakan Conditional Formatting untuk identifikasi anomali data
.
Membuat Sparklines untuk melihat tren performa keuangan secara ringkas
.
Menerapkan Data Validation dengan Error Alerts untuk menjaga integritas data
.
C4 (Analyze) - Menganalisis efektivitas pelaporan dan audit:
Menggunakan Slicers sebagai alat bantu filter visual yang interaktif.
Menerapkan Freeze Panes untuk manajemen tampilan data berskala besar.

--------------------------------------------------------------------------------
## 🎯 MATERI PEMBELAJARAN TERSTRUKTUR
### Blok 1: Visual Auditing (Conditional Formatting)
Definisi: Fitur untuk mengubah format sel secara otomatis berdasarkan kriteria tertentu untuk mempermudah pemeriksaan (auditing)
.
Highlight Cells Rules: Menandai saldo minus (merah) atau pendapatan di atas target (hijau).
Duplicate Values: Menemukan entri ganda yang terlewat pada tahap pembersihan data di pertemuan sebelumnya
.

### Blok 2: Micro Charts & Visual Filtering (Sparklines & Slicers)
Definisi: Alat bantu untuk menyajikan tren dan navigasi data secara visual tanpa grafik yang kompleks.
Sparklines: Grafik mini dalam satu sel (Line atau Column) untuk melihat fluktuasi kas/penjualan bulanan
.
Slicers: Tombol filter visual yang terhubung dengan Table (mempermudah user non-teknis melakukan filter data).

### Blok 3: Data Integrity & View Management (Validation & Freeze Panes)
Definisi: Memastikan data yang dimasukkan benar sejak awal dan mempermudah navigasi laporan yang panjang.
Data Validation (Advanced): Mengatur pesan kesalahan (Error Alert) jika user memasukkan angka di luar batas yang ditentukan atau teks yang salah
.
Freeze Panes: Mengunci baris judul agar tetap terlihat saat menggulir (scrolling) ribuan data transaksi.

--------------------------------------------------------------------------------
## 📊 STRUKTUR DATA PRAKTIK
### Kasus Bisnis: Laporan Evaluasi Kinerja CV Komputer Indonesia
Konteks Bisnis: CV Komputer Indonesia sedang mengevaluasi performa penjualan dan arus kas di berbagai cabang selama 6 bulan terakhir
. Anda diminta untuk membuat dashboard sederhana yang bisa mendeteksi kesalahan input (audit) dan menunjukkan tren penjualan.

--------------------------------------------------------------------------------
## 🎯 LATIHAN PRAKTIKUM TERSTRUKTUR (2 LEVEL - 90 Menit)
### LEVEL 1: Visual Auditing & Trend Monitoring
Fokus: Conditional Formatting & Sparklines - Durasi: 40 menit
Instruksi:
Buka dataset "Penjualan Cabang".
Gunakan Conditional Formatting untuk memberi warna merah pada angka penjualan yang di bawah 50 juta.
Gunakan fitur Highlight Duplicates pada kolom "ID Transaksi" untuk memastikan tidak ada input ganda.
Pada kolom "Tren", buatlah Sparklines (Line) yang merangkum data penjualan dari bulan Januari hingga Juni.
Pertanyaan Analisis:
Berdasarkan Sparklines, cabang mana yang menunjukkan tren penurunan konsisten?

--------------------------------------------------------------------------------
### LEVEL 2: Data Integrity & Interactive Dashboard
Fokus: Data Validation, Slicers, Freeze Panes - Durasi: 50 menit
Instruksi:
Aktifkan Freeze Panes pada baris judul agar navigasi data tetap nyaman.
Terapkan Data Validation pada kolom "Target":
Hanya boleh diisi angka antara 0 - 100.
Buat Error Alert berjudul "Input Salah" dengan pesan "Nilai target tidak boleh lebih dari 100%".
Karena data sudah dalam bentuk Table, tambahkan Slicer berdasarkan "Wilayah" dan "Kategori Produk".
Gunakan Slicer untuk menampilkan hanya wilayah "Jawa Timur".
Validasi Akhir:
Cek apakah Slicer berfungsi memotong data secara otomatis.
Tes validasi data dengan memasukkan angka 150; pastikan pesan kesalahan muncul.

--------------------------------------------------------------------------------
## 💡 Tips & Best Practices
Gunakan Warna yang Intuitif: Hijau untuk positif/aman, Merah untuk negatif/bahaya (Standar Akuntansi)
.
Slicer untuk Presentasi: Saat mempresentasikan data ke manajemen, gunakan Slicer daripada filter manual di judul kolom karena lebih user-friendly.
Hati-hati dengan Conditional Formatting: Jangan terlalu banyak memberikan warna berbeda dalam satu sheet karena akan membuat laporan sulit dibaca (distorsi visual).

--------------------------------------------------------------------------------
## 📚 Checklist Penilaian
[ ] Berhasil menandai saldo anomali dengan warna otomatis.
[ ] Sparklines muncul dan mencerminkan tren data yang benar.
[ ] Pesan Error Alert muncul saat terjadi kesalahan input data.
[ ] Slicer berfungsi untuk menyaring data secara visual.
[ ] Judul kolom tetap terlihat saat data di-scroll ke bawah (Freeze Panes).