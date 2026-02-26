# 📊 Sistem Absensi Karyawan

> Sistem otomatis profesional untuk pengolahan data absensi karyawan dengan antarmuka modern dan user-friendly.

[![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/status-production-brightgreen.svg)]()

---

## 📚 Quick Links

> 📑 **[INDEX.md](INDEX.md)** - Navigasi lengkap semua dokumentasi

| Dokumen | Deskripsi |
|---------|-----------|
| [📖 PANDUAN.md](PANDUAN.md) | Panduan lengkap penggunaan (step-by-step) |
| [🎨 DEMO.md](DEMO.md) | Visual demo & preview tampilan aplikasi |
| [📕 TECHNICAL.md](TECHNICAL.md) | Dokumentasi teknis & arsitektur sistem |
| [🗺️ ROADMAP.md](ROADMAP.md) | Roadmap & contributing guidelines |
| [📄 SUMMARY.md](SUMMARY.md) | Ringkasan fitur & achievement |

---

## ⚡ Quick Start (3 Langkah)

### 1️⃣ Setup (Pertama kali saja)
```bash
# Linux/Mac
./setup.sh

# Windows
setup.bat
```

### 2️⃣ Jalankan
```bash
python run.py
```

### 3️⃣ Pilih Mode
- **Mode 1**: GUI (Grafis) - Recommended 🎨
- **Mode 2**: CLI (Terminal) - Untuk server/headless 💻

---

## ✨ Fitur Unggulan

### 🎯 Smart & Automatic
- ✅ **Auto-detect** format CSV (`;` atau `,`)
- ✅ **Auto-detect** format tanggal (DD-MM-YYYY, YYYY-MM-DD, dll)
- ✅ **Auto-populate** daftar karyawan dari CSV
- ✅ **Auto-generate** nama file output
- ✅ **Auto-calculate** durasi kerja
- ✅ **Auto-detect** hari libur nasional Indonesia 2026

### 🌐 Flexible & Universal
- ✅ Dual-mode interface (GUI + CLI)
- ✅ Cross-platform (Windows, Linux, macOS)
- ✅ Support multiple CSV separators
- ✅ Support multiple date formats
- ✅ Berfungsi dengan atau tanpa Tkinter

### 🇮🇩 Localized for Indonesia
- ✅ Nama hari dalam Bahasa Indonesia
- ✅ Nama bulan dalam Bahasa Indonesia
- ✅ Kalender libur nasional Indonesia
- ✅ Format laporan standar Indonesia

---

## 📦 Apa Yang Termasuk?

```
📂 Project Files
├── 🚀 run.py                      ⭐ START HERE!
├── 🎨 absensi_app.py             (GUI Mode)
├── 💻 absensi_cli.py             (CLI Mode)
├── ⚙️ attendance_processor.py   (Core Engine)
├── 🔧 setup.sh / setup.bat       (Auto Setup)
├── 📖 README.md                   (Quick Start)
├── 📘 PANDUAN.md                  (Complete Guide)
└── 📕 TECHNICAL.md                (Developer Docs)
```

---

## 📋 Persyaratan Sistem

- **Python**: 3.7 atau lebih baru
- **Dependencies**: pandas, openpyxl (auto-install via setup)
- **OS**: Windows, Linux, macOS

---

## 🚀 Instalasi Lengkap

### Opsi 1: Automatic Setup (Recommended)

**Linux/Mac:**
```bash
chmod +x setup.sh
./setup.sh
python run.py
```

**Windows:**
```cmd
setup.bat
python run.py
```

### Opsi 2: Manual Setup

```bash
# 1. Install dependencies
pip install pandas openpyxl

# 2. Jalankan aplikasi
python run.py
```

---

## 💻 Cara Menggunakan

### Mode 1: GUI (Grafis) - RECOMMENDED ✨

**Menjalankan:**
```bash
python absensi_app.py
```

**Interface:**
```
┌─────────────────────────────────────────────┐
│  📊 SISTEM ABSENSI KARYAWAN                 │
├─────────────────────────────────────────────┤
│  1️⃣ Pilih File Data Absensi                │
│  2️⃣ Pilih Karyawan                         │
│  3️⃣ Pilih Periode Laporan                  │
│  4️⃣ Nama File Output                       │
│  5️⃣ [▶️ PROSES DATA]                       │
└─────────────────────────────────────────────┘
```

**Langkah:**
1. Klik **"Browse"** → Pilih file CSV
2. Pilih nama karyawan dari dropdown
3. Pilih bulan dan tahun
4. Klik **"Auto"** atau ketik nama file
5. Klik **"PROSES DATA"**
6. ✅ File Excel tersimpan!

### Mode 2: CLI (Command Line) 💻

**Menjalankan:**
```bash
python absensi_cli.py
```

**Fitur:**
- Menu interaktif
- Pencarian karyawan
- Pagination (20 karyawan/halaman)
- Progress indicator
- Summary statistics

### Metode 2: Menggunakan Command Line

Edit file `laporan.py` untuk konfigurasi:

```python
NAMA_TARGET = "Nama Karyawan"  # Sesuaikan dengan nama karyawan
FILE_SUMBER = 'nama_file.csv'  # File sumber data
OUTPUT_FILE = 'output.xlsx'    # Nama file output
```

Jalankan:
```bash
python laporan.py
```

## 📁 Struktur File

```
laporan-excel/
├── absensi_app.py           # Aplikasi GUI utama
├── attendance_processor.py  # Core processor untuk pengolahan data
├── laporan.py              # Script command line (legacy)
├── requirements.txt        # Dependencies
├── README.md              # Dokumentasi
└── contoh/
    ├── absen-mentahan.csv
    ├── absen-mentahan2.csv
    └── absensi_kehadiran_bulan_januari.xlsx
```

## 📝 Format File Input

File CSV harus memiliki struktur:
- Kolom pertama: `nama` (nama karyawan)
- Kolom kedua: `nik` (nomor induk karyawan)
- Kolom selanjutnya: tanggal dengan format `DD-MM-YYYY`

Contoh:
```csv
nama;nik;01-01-2026;02-01-2026;03-01-2026
John Doe;12345;07:00:00 - 16:00:00;07:30:00 - 16:30:00;
```

## 📊 Format Output

File Excel yang dihasilkan berisi kolom:
- **Tanggal** - Tanggal absensi
- **Hari** - Nama hari (Senin, Selasa, dst)
- **Jam Masuk** - Waktu masuk
- **Jam Pulang** - Waktu pulang
- **Durasi Kerja** - Total jam kerja
- **Keterangan** - Status kehadiran (Hadir, Libur, Tidak Hadir, dll)

## 🏖️ Hari Libur Nasional 2026

Sistem otomatis mendeteksi hari libur nasional Indonesia:
- Tahun Baru (1 Januari)
- Isra Miraj (23 Maret)
- Nyepi (31 Maret)
- Wafat Yesus Kristus (3 April)
- Paskah (5 April)
- Hari Buruh (1 Mei)
- Kenaikan Yesus Kristus (4 Mei)
- Waisak (14 Mei)
- Hari Lahir Pancasila (1 Juni)
- Idul Fitri (17-18 Juni)
- Hari Kemerdekaan RI (17 Agustus)
- Idul Adha (24 Agustus)
- Tahun Baru Islam (14 September)
- Maulid Nabi Muhammad SAW (23 November)
- Natal (25 Desember)

## 🛠️ Troubleshooting

### Error: Module pandas not found
```bash
pip install pandas openpyxl
```

### Error: Kolom 'nama' tidak ditemukan
Pastikan file CSV menggunakan separator yang benar (`;` atau `,`) dan memiliki kolom `nama`

### File Excel tidak bisa dibuka
Pastikan tidak ada file dengan nama yang sama sedang terbuka di Excel

## 🔄 Update Hari Libur

Untuk menambahkan hari libur tahun lain, edit file `attendance_processor.py`:

```python
INDONESIAN_HOLIDAYS = {
    '2027': {
        '01-01': 'Tahun Baru 2027',
        # tambahkan libur lainnya
    }
}
```

## 📞 Support

Jika menemukan bug atau ada saran, silakan buat issue atau hubungi developer.

## 📄 License

MIT License - Free to use and modify

## 👨‍💻 Developer

Sistem Absensi Karyawan v1.0
Dikembangkan untuk kemudahan pengelolaan data absensi

---

## 🔧 Yang Sudah Diperbaiki

### ✅ Kesalahan yang Ditemukan & Solusi:

1. **Kurangnya Validasi Data**
   - ❌ **Masalah:** Script langsung proses tanpa cek kolom 'nama' dan kolom tanggal
   - ✅ **Solusi:** Tambah validasi untuk memastikan kolom ada sebelum diproses

2. **Error Handling Lemah**
   - ❌ **Masalah:** Tidak ada informasi detail saat error
   - ✅ **Solusi:** Tambah logging yang jelas dan informasi kolom yang tersedia

3. **Logika Status Hari Kurang Jelas**
   - ❌ **Masalah:** Sabtu/Minggu ditandai "Sabtu"/"Minggu", tidak konsisten
   - ✅ **Solusi:** Ubah menjadi "Libur" untuk weekend tanpa data absensi

4. **Hardcoded Values**
   - ❌ **Masalah:** Jabatan dan posisi cell di-hardcode
   - ✅ **Solusi:** Tambah komentar untuk memudahkan kustomisasi

5. **Pesan Log Kurang Informatif**
   - ❌ **Masalah:** User tidak tahu progress script
   - ✅ **Solusi:** Tambah header, progress bar konsol, dan pesan sukses/gagal yang jelas

---

## 📋 Cara Penggunaan

### 1. Persiapan File
Pastikan Anda punya 2 file di folder `laporan-excel`:
- `ekspor_csv.xlsx` - File data absensi dari sistem
- `absensi_kehadiran_bulan.xlsx` - File template yang akan diisi

### 2. Konfigurasi (Baris 7-12 di `laporan.py`)
```python
INPUT_DATA_FILE = 'ekspor_csv.xlsx'           # File Data Absensi
INPUT_TEMPLATE_FILE = 'absensi_kehadiran_bulan.xlsx' # File Template
OUTPUT_FILE = 'Laporan_Absensi_November_2025.xlsx'  # Output

TARGET_NAME = "I Made Brahmanda Setyadi, S.Kom"  # Nama pegawai
TARGET_MONTH_PREFIX = "2025-11" # Bulan target (format: YYYY-MM)
```

**Ubah sesuai kebutuhan Anda:**
- `TARGET_NAME` → Nama pegawai yang akan diproses
- `TARGET_MONTH_PREFIX` → Bulan yang ingin di-generate (misal: "2025-12")
- `OUTPUT_FILE` → Nama file output

### 3. Jalankan Script
```bash
python laporan.py
```

### 4. Hasil
File Excel baru akan dibuat dengan nama yang sudah Anda tentukan di `OUTPUT_FILE`.

---

## 📁 Struktur File Excel

### File Input (`ekspor_csv.xlsx`)
Harus punya struktur seperti ini:
```
| nama                              | 2025-11-01  | 2025-11-02  | 2025-11-03  | ...
|-----------------------------------|-------------|-------------|-------------|
| I Made Brahmanda Setyadi, S.Kom   | 08:00 - 17:00 | 08:15 - 17:10 |           |
```

### File Template (`absensi_kehadiran_bulan.xlsx`)
Template dengan header di baris tertentu, data mulai baris 14:
- Kolom B: Tanggal
- Kolom C: Jam Masuk
- Kolom D: Jam Pulang
- Kolom E: Durasi Kerja
- Kolom F: Keterangan/Status

---

## 🎨 Fitur Script

✅ **Auto-detect format file** (Excel atau CSV)  
✅ **Validasi kolom nama dan tanggal**  
✅ **Hitung durasi kerja otomatis**  
✅ **Deteksi hari libur (Sabtu/Minggu)**  
✅ **Hapus data lama sebelum isi yang baru**  
✅ **Format border dan alignment otomatis**  
✅ **Logging detail untuk debugging**  

---

## ⚠️ Troubleshooting

### Error: "Kolom 'nama' tidak ditemukan"
➡️ Pastikan file `ekspor_csv.xlsx` punya kolom dengan header **'nama'** (huruf kecil semua)

### Error: "Tidak ada kolom tanggal dengan prefix..."
➡️ Cek kolom tanggal di file Excel, pastikan formatnya `YYYY-MM-DD` (misal: 2025-11-01)

### Error: "File sedang dibuka di Excel"
➡️ Tutup file Excel output, lalu jalankan script lagi

### Data tidak muncul
➡️ Cek nilai `START_ROW` di line 107 (default: 14), sesuaikan dengan template Anda

---

## 🔄 Kustomisasi Lanjutan

### Mengubah Baris Awal Data
Edit line 107:
```python
START_ROW = 14  # Ubah sesuai template Anda
```

### Mengubah Posisi Header
Edit line 110-112:
```python
ws['F8'] = datetime.now().strftime('%Y-%m-%d')  # Cell tanggal cetak
ws['F10'] = TARGET_NAME                          # Cell nama pegawai
ws['F11'] = "Full Stack Web Developer"          # Cell jabatan
```

### Menambah Logika Status Custom
Edit function `calculate_duration_and_status()` di line 17-34.

---

## 📝 Catatan Penting

1. **Backup file template** Anda sebelum run script pertama kali
2. **Format tanggal** di ekspor harus konsisten (YYYY-MM-DD)
3. **Format jam** harus `HH:MM:SS - HH:MM:SS` (misal: 08:00:00 - 17:00:00)
4. Script akan **overwrite** data lama di template (baris 14-50)

---

## 🚀 Contoh Output Console

```
============================================================
  PEMBUATAN LAPORAN ABSENSI OTOMATIS
============================================================
[INFO] File Input   : ekspor_csv.xlsx
[INFO] Template     : absensi_kehadiran_bulan.xlsx
[INFO] File Output  : Laporan_Absensi_November_2025.xlsx
[INFO] Target Bulan : 2025-11
[INFO] Target Nama  : I Made Brahmanda Setyadi, S.Kom
============================================================

[OK] Berhasil membaca ekspor_csv.xlsx sebagai Excel (.xlsx)
[OK] Data pegawai 'I Made Brahmanda Setyadi, S.Kom' ditemukan.
[OK] Template dimuat.
   [...] Membersihkan data lama dari baris 14 sampai 50...
[OK] Ditemukan 30 hari data absensi untuk bulan 2025-11
[OK] Berhasil menulis 30 baris data absensi.

============================================================
[SUKSES] Laporan Absensi berhasil dibuat!
[INFO] File tersimpan di: Laporan_Absensi_November_2025.xlsx
============================================================
```

---

## 👨‍💻 Dikembangkan untuk
Otomasi pembuatan laporan absensi bulanan pegawai dengan data dari sistem ekspor Excel.

**Last Updated:** 1 Desember 2025
