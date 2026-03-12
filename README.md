# BRI Rekening Koran → Excel Web App

Upload PDF rekening koran BRI, otomatis diproses menjadi file Excel dengan:
- Rekap per bulan (debet, kredit, penjualan)
- Kategorisasi otomatis: Penjualan vs Non penjualan
- Sheet "Edit Penjualan" dengan dropdown untuk edit manual
- Summary otomatis terupdate via formula Excel

## Deploy ke Railway (gratis)

1. Push repo ini ke GitHub
2. Buka [railway.app](https://railway.app) → New Project → Deploy from GitHub
3. Pilih repo ini → Railway otomatis detect Python dan deploy
4. Selesai! Dapat URL publik dalam ~2 menit

## Jalankan lokal

```bash
pip install -r requirements.txt
python app.py
```

Buka http://localhost:5000
