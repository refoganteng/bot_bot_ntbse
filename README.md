# bot_bot_ntbse

Bot Otomasi Input Data dan Auto-Submit FASIH BPS (Sensus Ekonomi 2026).

## Fitur Utama
- **Login Manual Aman**: Browser Google Chrome asli dibuka otomatis, user login manual sehingga tidak memicu deteksi bot Cloudflare/WAF.
- **Deteksi Otomatis & Presisi**:
  - Penyesuaian tahun beroperasi (<= 2025 vs >= 2026).
  - Mengisi rincian biaya produksi, operasional, pendapatan, dan catatan/remark `#DC_01`.
  - **Auto-Resolve Galat Anomali**: Otomatis mendeteksi galat anomali pada modal kirim, membuka rincian galat, dan mengaktifkan toggle switch anomali.
- **Fail-Fast & Auto Skip**: Jika ada input yang tidak ditemukan atau stuck melebihi batas waktu, baris langsung di-skip dan tetap berstatus pending di Excel untuk dieksekusi ulang.
- **Anti-Sleep PC**: Bekerja otomatis di Windows (via Win32 API) dan macOS (via `caffeinate`) agar PC tidak sleep selama bot berjalan.
- **Auto Maximize Window**: Menyesuaikan resolusi dan rasio browser ke layar monitor secara optimal via Chrome DevTools Protocol (CDP).
- **Multi-OS / Cross-Platform**: Mendukung Windows, macOS, dan Linux tanpa konfigurasi manual path Chrome.

## Persyaratan
- Python 3.9+
- Google Chrome
- Sistem Operasi: Windows 10/11, macOS, atau Linux

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
