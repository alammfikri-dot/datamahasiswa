# SIM Mahasiswa — Flask + Python

Sistem Informasi Manajemen Mahasiswa berbasis Python Flask.

## Fitur
- Login admin
- Dashboard statistik (total mahasiswa, rata-rata IPK, IPK tertinggi, distribusi jurusan)
- CRUD data mahasiswa
- Sorting (nama / IPK)
- Pencarian algoritmik: Linear, Sequential, Binary Search, Rentang IPK
- Kirim email evaluasi akademik via Gmail SMTP
- Export data ke Excel (.xlsx)

## Cara Menjalankan

### 1. Install dependencies
```
pip install -r requirements.txt
```

### 2. Isi kredensial di file `.env`
Buka file `.env`, ganti baris ini:
```
GMAIL_APP_PASS = ISI_APP_PASSWORD_DI_SINI
```
dengan App Password Gmail kamu (bukan password biasa).

Cara dapat App Password Gmail:
1. Buka https://myaccount.google.com/security
2. Aktifkan 2-Step Verification
3. Cari "App passwords" → buat baru → pilih "Mail"
4. Salin 16 karakter yang muncul → paste ke .env

### 3. (Opsional) Taruh gambar background
Simpan file gambar bernama `3.jpg` ke folder `static/`

### 4. Jalankan server
```
python app.py
```

### 5. Buka browser
```
http://localhost:5000
```

Login: `admin` / `admin123`

## Struktur Folder
```
sim_mahasiswa/
├── app.py              ← Backend utama (Flask + SMTP)
├── .env                ← Kredensial (JANGAN di-share/commit)
├── .gitignore
├── requirements.txt
├── README.md
├── templates/
│   ├── login.html
│   └── index.html
└── static/
    └── 3.jpg           ← Background login (taruh sendiri)
```
