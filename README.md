# 🕌 Portal Pendidikan Bunayya Islamic School

Sistem manajemen pendidikan berbasis web untuk Bunayya Islamic School — mencakup Portal Guru, Portal Siswa, dan Portal Wali.

## ✨ Fitur

- 📈 Dashboard statistik dengan grafik bar chart
- 📋 Absensi Siswa & Guru
- 📝 Jurnal Mengajar
- 📊 Rekap Nilai per Kelas
- 🎯 Laporan Kegiatan & Target Pencapaian
- 📖 Ziyadah Hafalan
- 🏡 Portal Wali Murid
- ☁️ Sinkronisasi ke Google Sheets

## 🚀 Cara Akses

Buka langsung: [bunayya-portal.vercel.app](https://bunayya-portal.vercel.app)

## 🔑 Login Demo

| Portal | Username | Password |
|--------|----------|----------|
| Admin  | admin    | admin123 |
| Guru   | (pilih nama guru) | 1234 |
| Wali   | (pilih kelas & siswa) | — |

## 📁 File

| File | Keterangan |
|------|-----------|
| `index.html` | Aplikasi portal lengkap |
| `kode.gs` | Google Apps Script untuk sinkronisasi Google Sheets |
| `vercel.json` | Konfigurasi deployment Vercel |

## ⚙️ Setup Google Sheets

1. Buka Google Sheet ID: `1NyNLjFhDOqsi4KYGM-lmB3ZGupt7amn2W67kYUIac-Y`
2. Extensions → Apps Script → paste isi `kode.gs`
3. Deploy sebagai Web App → salin URL
4. Tempel URL ke variabel `SCRIPT_URL` di `index.html`

---
© 2025 Bunayya Islamic School
