# 📌 Project System Automasi SSO to Summary

**Project System Automasi SSO to Summary** adalah sebuah sistem otomatisasi yang memindahkan data dari hasil output SSO ke dalam file Summary. Sistem ini dirancang untuk mempermudah proses penyusunan serta pengelolaan database Summary pada tim CBIC, sehingga alur kerja menjadi lebih efisien, konsisten, dan terstruktur.

---

## ✨ Fitur Utama
- ✅ Mengolah data dari file Excel (`data/`)
- ✅ GUI sederhana untuk interaksi pengguna (`gui/`)
- ✅ Struktur project modular dengan folder `app/`, `logic/`, dan `style/`
- ✅ Dukungan untuk Windows (`run.bat`) dan Linux/Mac (`setup.sh`)

---
## 📂 Struktur Folder

```
project-name/
│── app/ # Core aplikasi
│── assets/ # Gambar, ikon, atau asset lainnya
│── config/ # File konfigurasi JSON / settings
│── gui/ # Modul GUI (interface)
│── logic/ # Modul logika bisnis / processing
│── style/ # File style / tema UI
│── main.py # Entry point aplikasi
│── run.bat # Script untuk menjalankan aplikasi di Windows
│
├── data/ # Dataset / file input
│ ├── Sample_Input.xlsx
│
├── .gitignore # File gitignore
├── README.md # Dokumentasi
└── requirements.txt # Dependensi Python
```

## ⚙️ Instalasi & Menjalankan

### 1. Install Python
Pastikan Python **3.8 atau lebih baru** sudah terinstall.  
Download di: [Python.org](https://www.python.org/downloads/)

Cek apakah Python sudah terinstall dengan:
```bash
python --version

```

## ⚙️ Clone Repository

1. Clone repository ini:
```
git clone https://github.com/ITM-CBIC-Team/SystemAutomation_SSOtoSummary.git
```

## Install Requirement yang dibutuhkan

```
pip install -r requirements.txt
```

## Jalankan run.bat
```
run.bat
```