# Aplikasi KeuanganKu

Aplikasi pencatatan keuangan pribadi berbasis Google Apps Script dengan Google Sheets sebagai database.

## Fitur
- 📊 **Dashboard** — Ringkasan saldo, pemasukan, pengeluaran, grafik arus kas
- 💰 **Catat Transaksi** — Pemasukan, pengeluaran, transfer antar dompet, import CSV
- 📋 **Riwayat Transaksi** — Buku besar lengkap dengan filter dan pencarian
- 🎯 **Anggaran** — Pengelolaan anggaran bulanan per kategori
- 📈 **Investasi** — Pencatatan portofolio investasi
- ⚙️ **Pengaturan** — Manajemen dompet dan kategori

## Teknologi
- Google Apps Script (backend)
- Google Sheets (database)
- Tailwind CSS + Font Awesome (UI)
- Chart.js (grafik)

## Deployment
App ini di-deploy sebagai Google Apps Script Web App.  
URL: `https://script.google.com/macros/s/AKfycbzwEYqyjw04RtD1MquvdFVv6IgQSyiKyzO7y_m_SPTeGvzztwf9GtCkvD1ghMl6CeqN/exec`

## Development
Menggunakan [clasp](https://github.com/google/clasp) untuk sinkronisasi kode.

```bash
# Install clasp
npm install -g @google/clasp

# Login
clasp login

# Pull kode terbaru dari GAS
clasp pull

# Push perubahan ke GAS
clasp push

# Deploy ulang
clasp deploy
```
