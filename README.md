# RATIG Desktop (Reaction Timer & Fatigue Detection System)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![GUI](https://img.shields.io/badge/GUI-Tkinter-orange)
![Database](https://img.shields.io/badge/Database-SQLite3-lightgrey?logo=sqlite)
![Hardware](https://img.shields.io/badge/Hardware-ESP32%20Serial-green)
![Build](https://img.shields.io/badge/Installer-PyInstaller%20%7C%20Inno%20Setup-red)

Aplikasi Desktop Monitoring Kinerja Refleks dan Tingkat Kelelahan (Fatigue) Pekerja secara Real-Time menggunakan Pengujian Reaction Timer Berbasis Hardware ESP32.

---

## 🌟 Fitur Utama

- **Real-Time Serial Communication**:
  - Koneksi otomatis dan deteksi port COM serial ke alat Reaction Timer ESP32.
  - Pembacaan data waktu reaksi (milidetik/ms) secara langsung dari pengujian fisik.
- **Manajemen Data Master Karyawan & Kontraktor**:
  - Pengelolaan database data pekerja, departemen, dan shift kerja.
  - Impor dan ekspor data master via file Excel (`Master_Karyawan.xlsx` / `Master_Kontraktor.xlsx`).
- **Analisis Tingkat Kelelahan (Fatigue Analytics)**:
  - Algoritma evaluasi refleks dan tingkat fatigue (Normal, Kelelahan Ringan, Kelelahan Sedang, Kelelahan Berat).
  - Integrasi Grafik Statistik Visualisasi Tren menggunakan **Matplotlib**.
- **Pelaporan & Ekspor Data**:
  - Ekspor hasil pengujian otomatis ke format Excel (`.xlsx`) dan PDF Report (`.pdf`).
  - Fitur sistem backup dan restore database SQLite secara otomatis (`backups/`).
- **Antarmuka & Pengalaman Pengguna (GUI)**:
  - Antarmuka berbasis Tkinter modern dengan tema visual intuitif.
  - Dukungan pencetakan skematik PCB / tata letak alat (`PCB_ESP32_Printable.html`).

---

## 📁 Struktur Folder Project

```text
RATIG_Desktop/
├── RATIG.py                    # Script Utama Aplikasi Desktop (Tkinter GUI & Logic)
├── RATIG.spec                  # Konfigurasi PyInstaller Executable Build
├── build.bat                   # Batch script untuk otomatisasi kompilasi ke EXE
├── build_installer.iss         # Script Inno Setup untuk membuat file Installer Windows
├── PCB_ESP32_Printable.html    # Layout & Dokumentasi Cetak PCB Hardware ESP32
├── requirements.txt            # Daftar dependensi Python
├── config/                     # Konfigurasi aplikasi
│   └── settings.json
├── assets/                     # Gambar, logo, & ikon (.ico)
│   └── ratig.ico
└── README.md
```

---

## 🛠️ Persyaratan Sistem

- **Sistem Operasi**: Windows 10 / 11 (32-bit & 64-bit)
- **Python**: Versi 3.8 atau lebih baru
- **Hardware Perangkat**: Perangkat Test Reaction Timer ESP32 dengan koneksi USB Serial (CH340 / CP210x / FTDI)

---

## 🚀 Cara Menjalankan Aplikasi

### 1. Install Dependensi Python
Buka terminal / command prompt pada folder project dan jalankan:
```bash
pip install -r requirements.txt
```

### 2. Jalankan Script Python
```bash
python RATIG.py
```

---

## 📦 Menghasilkan File Executable (`.exe`) & Installer

### 1. Build Single Portable Executable (`RATIG.exe`)
Jalankan batch script build yang telah disediakan:
```cmd
build.bat
```
Atau jalankan PyInstaller secara manual:
```bash
pyinstaller RATIG.spec
```
*Hasil file `.exe` akan berada di folder `dist/` atau `release/RATIG_Portable/`.*

### 2. Membuat Installer Windows (`.exe` Setup)
Buka file `build_installer.iss` menggunakan aplikasi **Inno Setup Compiler**, lalu klik **Compile**.

---

## 📝 Lisensi
Project ini dikembangkan untuk kebutuhan monitoring keselamatan kerja dan fatigue manajemen di lingkungan kerja.
