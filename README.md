# 🚀 Otomasi Summary Event Promo

Aplikasi web untuk mengotomasi pembuatan summary dari data event promo dalam format Excel.

## 📌 Fitur

- ✅ Upload file Excel dengan multiple sheets
- ✅ Otomatis mengekstrak informasi promo dari setiap sheet
- ✅ Menghasilkan summary dalam format Excel yang rapi
- ✅ Preview hasil sebelum download
- ✅ Format angka dengan separator ribuan
- ✅ NaN ditampilkan sebagai kosong

## 🎯 Format Input yang Diharapkan

File Excel dengan struktur:
- Multiple sheets (setiap sheet = 1 promo)
- **Row 2**: Header dengan nama promo dan periode
- **Row 3**: Mekanisme promo
- **Row 7**: Data summary (total)

## 📊 Kolom Output

| Kolom | Keterangan |
|-------|------------|
| No. | Nomor urut |
| Nama Promo | Nama program promo (tanpa tanggal) |
| Mekanisme Promo | Detail mekanisme/syarat promo |
| Periode Promo | Tanggal berlaku promo |
| All Count | Total customer |
| All Claim | Total claim |
| Sales Amount | Total nilai penjualan |
| Amount | Total nilai bonus/hadiah |
| Left | Sisa bonus/hadiah |

## 🚀 Cara Menggunakan

1. Buka aplikasi di browser
2. Upload file Excel mentah (.xlsx)
3. Klik tombol **Proses File**
4. Lihat preview hasil
5. Klik **Download Summary Excel**

## 🛠️ Tech Stack

- Python 3.9+
- Streamlit
- Pandas
- OpenPyXL

## 📝 Version

**v1.0** - January 2026

