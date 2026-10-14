# bot_bot_ntbse

Bot Otomasi Input Data dan Auto-Submit FASIH BPS (Sensus Ekonomi 2026).

## Fitur Utama
- **Login Manual Aman**: Browser Google Chrome asli dibuka otomatis, user login manual sehingga tidak memicu deteksi bot Cloudflare/WAF.
- **Deteksi Otomatis & Presisi**:
  - Penyesuaian tahun beroperasi (<= 2025 vs >= 2026).
  - Mengisi rincian biaya produksi, operasional, pendapatan, dan catatan/remark `#DC_01`.
  - **Auto-Resolve Galat Anomali**: Otomatis mendeteksi galat anomali pada modal kirim, membuka rincian galat, dan mengaktifkan toggle switch anomali.
- **Fail-Fast & Auto Skip**: Jika ada input yang tidak ditemukan atau stuck melebihi batas waktu, baris langsung di-skip dan tetap berstatus pending di Excel untuk dieksekusi ulang.
- **Anti-Sleep PC**: Memanfaatkan utilitas macOS `caffeinate` agar Mac tidak masuk mode sleep selama bot berjalan.
- **Auto Maximize Window**: Menyesuaikan resolusi dan rasio browser ke layar monitor secara optimal.

## Persyaratan
- Python 3.9+
- Google Chrome
- macOS (didukung `caffeinate` dan AppleScript)

## Instalasi
```bash
pip install -r requirements.txt
python -m playwright install
```

## Cara Menjalankan
```bash
python auto_input_fasih.py
```
*(Opsional: `python auto_input_fasih.py --dry-run` untuk simulasi tanpa klik submit akhir)*
