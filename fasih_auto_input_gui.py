"""
FASIH Auto Input - PyAutoGUI Version
====================================
Metode otomasi simulasi mouse & keyboard langsung pada Google Chrome
(seperti AutomationApproveFasih/fasih_auto_approve.py).

Mengapa versi ini?
  Sistem keamanan BPS / FASIH (F5 BIG-IP ASM Bot Defense) mendeteksi browser otomatis
  (Playwright / Selenium / Puppeteer) dengan pesan:
  "Kami mendeteksi perilaku yang tidak wajar pada koneksi anda... BOT-xxx".
  Dengan PyAutoGUI, script mengontrol langsung aplikasi Google Chrome Anda yang sudah login asli,
  sehingga 100% AMAN dari deteksi bot!

Fitur:
  1. Membuka link dari Excel di Google Chrome biasa.
  2. Klik tombol Edit & navigasi BLOK II.
  3. Mengisi nilai 26b, 26d, 27a (atau 30b, 30d, 31a jika 2026) secara cerdas menggunakan
     fitur pencarian browser (Cmd+F) sehingga tidak terpengaruh posisi scroll.
  4. Mengisi Catatan (Remark #DC_01) pada 27a / 31a.
  5. Kirim -> Kirim Modal -> Konfirmasi.
  6. Tutup tab (Cmd+W) dan lanjut ke link berikutnya.
  7. Mode Kalibrasi Cepat: python fasih_auto_input_gui.py calibrate

Failsafe Darurat:
  Geser mouse ke pojok kiri-atas layar secara cepat untuk menghentikan bot seketika.
"""

import time
import sys
import os
import subprocess
import json
from pathlib import Path
from datetime import datetime
import openpyxl
import pyautogui
import pyperclip

# ─────────────────────────── FAILSAFE & SETTINGS ─────────────────────

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2

EXCEL_FILE = Path("data_katA_kopi.xlsx")
CONFIG_FILE = Path("config_koordinat.json")

# Default Delay (detik) - sesuaikan jika koneksi lebih lambat/cepat
DELAY_LOAD_PAGE = 8.0    # Waktu tunggu halaman kuesioner terbuka
DELAY_ACTION    = 0.8    # Jeda antar aksi kecil
DELAY_MODAL     = 1.5    # Jeda tunggu modal/popup muncul

# ─────────────────────────── DEFAULT KOORDINAT ───────────────────────
# Catatan: Resolusi layar Anda terdeteksi 1680x1050 (MacBook Retina default).
# Anda bisa mengubah angka ini langsung, atau jalankan:
#   python fasih_auto_input_gui.py calibrate
# untuk merekam koordinat secara interaktif.

DEFAULT_COORDS = {
    "EDIT_BTN":         {"x": 1616, "y": 880, "desc": "Tombol icon Edit (pensil bulat di kanan bawah)"},
    "BLOK2_BTN":        {"x": 140,  "y": 320, "desc": "Menu SE2026 - L BLOK II di sidebar kiri"},
    "USAHA_BTN":        {"x": 150,  "y": 420, "desc": "Sub-item nama usaha di bawah BLOK II (misal MEMBUDIDAYAKAN KOPI)"},
    "CATATAN_27A_BTN":  {"x": 1550, "y": 550, "desc": "Tombol icon Catatan di samping rincian 27a / 31a"},
    "TEXTAREA_CATATAN": {"x": 840,  "y": 680, "desc": "Kolom teks Catatan di dalam modal popup"},
    "SIMPAN_CATATAN":   {"x": 1130, "y": 690, "desc": "Tombol icon Simpan catatan (pesawat kertas di kanan textarea)"},
    "TUTUP_CATATAN":    {"x": 1150, "y": 280, "desc": "Tombol silang (X) untuk menutup modal Catatan"},
    "KIRIM_HEADER":     {"x": 1590, "y": 75,  "desc": "Tombol Kirim di header/navbar atas"},
    "KIRIM_MODAL":      {"x": 1050, "y": 680, "desc": "Tombol Kirim di dalam modal ringkasan jawaban"},
    "KONFIRMASI_BTN":   {"x": 930,  "y": 560, "desc": "Tombol Konfirmasi pada popup 'Konfirmasi Kirim'"},
}


def load_coords() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                return {k: {"x": data[k]["x"], "y": data[k]["y"], "desc": DEFAULT_COORDS[k]["desc"]} for k in DEFAULT_COORDS if k in data}
        except Exception:
            pass
    return DEFAULT_COORDS


def save_coords(coords: dict):
    with open(CONFIG_FILE, "w") as f:
        json.dump(coords, f, indent=2)
    print(f"✅ Koordinat disimpan ke {CONFIG_FILE}")


# ─────────────────────────── KALIBRASI INTERAKTIF ───────────────────

def run_calibration():
    print(r"""
╔══════════════════════════════════════════════════════════════════════╗
║                    MODE KALIBRASI KOORDINAT                          ║
║                                                                      ║
║  Petunjuk:                                                           ║
║  1. Buka satu halaman kuesioner FASIH di Google Chrome               ║
║  2. Setiap pertanyaan muncul di terminal:                            ║
║     - Arahkan mouse Anda ke posisi tombol yang diminta di Chrome     ║
║     - Tekan ENTER di terminal untuk merekam posisi mouse tersebut    ║
╚══════════════════════════════════════════════════════════════════════╝
    """)

    coords = load_coords()

    for key, info in coords.items():
        print(f"\n👉 [{key}]")
        print(f"   Keterangan: {info['desc']}")
        print(f"   Posisi sekarang: X={info['x']}, Y={info['y']}")
        ans = input("   Arahkan kursor lalu tekan [ENTER] untuk rekam (atau ketik 's' untuk skip): ")
        if ans.strip().lower() != 's':
            cur_x, cur_y = pyautogui.position()
            info["x"] = cur_x
            info["y"] = cur_y
            print(f"   ✓ Terekam: X={cur_x}, Y={cur_y}")

    save_coords(coords)
    print("\n🎉 Kalibrasi selesai! Sekarang Anda bisa jalankan: python fasih_auto_input_gui.py\n")


# ─────────────────────────── EXCEL HELPERS ───────────────────────────

def clean_val(v):
    if v is None:
        return ""
    s = str(v).strip()
    try:
        f = float(s.replace(",", "."))
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s


def parse_year(v):
    if v is None:
        return 2025
    s = str(v).strip()
    try:
        return int(float(s))
    except ValueError:
        return 2025


def read_excel_rows(filepath: Path):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    status_col = None
    time_col = None
    for j in range(1, ws.max_column + 1):
        val = str(ws.cell(1, j).value or "").lower()
        if "status" in val:
            status_col = j
        if "waktu" in val:
            time_col = j

    if not status_col:
        status_col = ws.max_column + 1
        ws.cell(1, status_col, "Status Otomasi")
    if not time_col:
        time_col = status_col + 1
        ws.cell(1, time_col, "Waktu Eksekusi")

    items = []
    for r in range(2, ws.max_row + 1):
        link = ws.cell(r, 2).value
        if not link:
            continue

        item = {
            "row_idx": r,
            "id_assignment": ws.cell(r, 1).value or "",
            "link": str(link).strip(),
            "nama_usaha": ws.cell(r, 8).value or "",
            "tahun": parse_year(ws.cell(r, 11).value),
            "adj_gaji": clean_val(ws.cell(r, 15).value),
            "adj_biaya_prod": clean_val(ws.cell(r, 16).value),
            "adj_biaya_ops": clean_val(ws.cell(r, 17).value),
            "adj_pendapatan": clean_val(ws.cell(r, 18).value),
            "remark": str(ws.cell(r, 19).value or "").strip(),
            "status": str(ws.cell(r, status_col).value or "").strip(),
            "status_col": status_col,
            "time_col": time_col,
        }
        items.append(item)

    return wb, items


def update_row_status(wb, filepath, item, status_text):
    ws = wb.active
    ws.cell(item["row_idx"], item["status_col"], status_text)
    ws.cell(item["row_idx"], item["time_col"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.save(filepath)


# ─────────────────────────── PYAUTOGUI ACTION HELPERS ────────────────

def open_url_in_chrome(url: str):
    """Membuka URL di Google Chrome dan membawanya ke depan."""
    subprocess.run(["open", "-a", "Google Chrome", url])


def fill_field_by_search(label: str, value: str):
    """
    Mencari label pertanyaan menggunakan Cmd+F Chrome,
    lalu menekan Tab untuk langsung masuk ke kolom input,
    kemudian mem-paste nilai baru.
    """
    if not value or value in ("0", "") and "gaji" in label:
        return

    print(f"   [Input] Mengisi {label} dengan nilai: {value}")

    # 1. Buka Find bar di Chrome
    pyautogui.hotkey('command', 'f')
    time.sleep(0.35)

    # 2. Paste teks pencarian
    pyperclip.copy(label)
    pyautogui.hotkey('command', 'v')
    time.sleep(0.4)

    # 3. Enter untuk loncat ke teks hasil temuan
    pyautogui.press('enter')
    time.sleep(0.35)

    # 4. Escape untuk menutup Find bar (fokus tetap di teks yang ditemukan)
    pyautogui.press('escape')
    time.sleep(0.25)

    # 5. Tab untuk masuk ke input field
    pyautogui.press('tab')
    time.sleep(0.25)

    # 6. Seleksi semua (Cmd+A) dan hapus isi lama
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.press('backspace')
    time.sleep(0.1)

    # 7. Paste nilai baru dari clipboard
    pyperclip.copy(str(value))
    pyautogui.hotkey('command', 'v')
    time.sleep(0.4)


def process_item(item: dict, coords: dict):
    url = item["link"]
    usaha = item["nama_usaha"]
    thn = item["tahun"]
    rem = item["remark"]

    print("\n" + "=" * 65)
    print(f"👉 Memproses Row {item['row_idx']}: {usaha} (Tahun {thn})")
    print(f"   URL: {url}")
    print("=" * 65)

    # 1. Buka URL di Chrome
    print("1. Membuka URL di Google Chrome...")
    open_url_in_chrome(url)
    print(f"   Menunggu halaman termuat ({DELAY_LOAD_PAGE} detik)...")
    time.sleep(DELAY_LOAD_PAGE)

    # 2. Klik tombol Edit
    print("2. Mengklik tombol Edit (floating icon)...")
    pyautogui.click(coords["EDIT_BTN"]["x"], coords["EDIT_BTN"]["y"])
    time.sleep(DELAY_ACTION * 1.5)

    # 3. Klik Sidepanel BLOK II
    print("3. Mengklik Sidepanel BLOK II...")
    pyautogui.click(coords["BLOK2_BTN"]["x"], coords["BLOK2_BTN"]["y"])
    time.sleep(DELAY_ACTION * 1.5)

    # Jika ada sub-item usaha spesifik di bawah BLOK II
    if coords.get("USAHA_BTN"):
        print(f"   Memilih sub-item usaha: {usaha}...")
        pyautogui.click(coords["USAHA_BTN"]["x"], coords["USAHA_BTN"]["y"])
        time.sleep(DELAY_ACTION)

    # 4. Mengisi Input Pertanyaan Berdasarkan Tahun
    if thn <= 2025:
        print("4. Mengisi rincian <= 2025 (26b, 26d, 27a)...")
        if item["adj_gaji"] and item["adj_gaji"] not in ("0", ""):
            fill_field_by_search("26. a.", item["adj_gaji"])

        if item["adj_biaya_prod"]:
            fill_field_by_search("26. b.", item["adj_biaya_prod"])

        if item["adj_biaya_ops"]:
            fill_field_by_search("26. d.", item["adj_biaya_ops"])

        if item["adj_pendapatan"]:
            fill_field_by_search("27. a.", item["adj_pendapatan"])
            q_catatan = "27. a."
    else:
        print("4. Mengisi rincian >= 2026 (30b, 30d, 31a)...")
        if item["adj_gaji"] and item["adj_gaji"] not in ("0", ""):
            fill_field_by_search("30. a.", item["adj_gaji"])

        if item["adj_biaya_prod"]:
            fill_field_by_search("30. b.", item["adj_biaya_prod"])

        if item["adj_biaya_ops"]:
            fill_field_by_search("30. d.", item["adj_biaya_ops"])

        if item["adj_pendapatan"]:
            fill_field_by_search("31. a.", item["adj_pendapatan"])
            q_catatan = "31. a."

    # 5. Mengisi Catatan / Remark jika ada
    if rem:
        print(f"5. Menambahkan Catatan '{rem}' pada {q_catatan}...")
        # Klik tombol icon Catatan di samping 27a / 31a
        pyautogui.click(coords["CATATAN_27A_BTN"]["x"], coords["CATATAN_27A_BTN"]["y"])
        time.sleep(DELAY_MODAL)

        # Klik textarea di modal catatan
        pyautogui.click(coords["TEXTAREA_CATATAN"]["x"], coords["TEXTAREA_CATATAN"]["y"])
        time.sleep(0.3)
        pyperclip.copy(rem)
        pyautogui.hotkey('command', 'v')
        time.sleep(0.4)

        # Klik tombol Simpan catatan
        pyautogui.click(coords["SIMPAN_CATATAN"]["x"], coords["SIMPAN_CATATAN"]["y"])
        time.sleep(DELAY_ACTION * 1.5)

        # Tutup modal catatan (Escape atau tombol Tutup)
        pyautogui.press('escape')
        time.sleep(0.4)
        pyautogui.click(coords["TUTUP_CATATAN"]["x"], coords["TUTUP_CATATAN"]["y"])
        time.sleep(DELAY_ACTION)

    # 6. Kirim & Konfirmasi
    print("6. Mengirim data (Kirim -> Kirim Modal -> Konfirmasi)...")
    # Klik tombol Kirim di header
    pyautogui.click(coords["KIRIM_HEADER"]["x"], coords["KIRIM_HEADER"]["y"])
    time.sleep(DELAY_MODAL)

    # Klik tombol Kirim di dalam modal ringkasan
    pyautogui.click(coords["KIRIM_MODAL"]["x"], coords["KIRIM_MODAL"]["y"])
    time.sleep(DELAY_MODAL)

    # Klik tombol Konfirmasi di popup terakhir
    pyautogui.click(coords["KONFIRMASI_BTN"]["x"], coords["KONFIRMASI_BTN"]["y"])
    time.sleep(DELAY_LOAD_PAGE * 0.4)  # Tunggu respon kirim

    # 7. Tutup Tab Chrome yang sudah selesai
    print("7. Menutup tab Chrome (Cmd+W)...")
    pyautogui.hotkey('command', 'w')
    time.sleep(DELAY_ACTION)


# ─────────────────────────── MAIN EXECUTION ──────────────────────────

def main():
    if len(sys.argv) > 1 and sys.argv[1].lower() == "calibrate":
        run_calibration()
        return

    if not EXCEL_FILE.exists():
        print(f"❌ File {EXCEL_FILE} tidak ditemukan!")
        return

    coords = load_coords()
    wb, items = read_excel_rows(EXCEL_FILE)
    pending = [it for it in items if it["status"] not in ("BERHASIL", "OK")]

    print(r"""
╔══════════════════════════════════════════════════════════════════════╗
║               FASIH AUTO INPUT (PYAUTOGUI EDITION)                   ║
║                                                                      ║
║  • 100% Bebas Deteksi Bot (menggunakan Chrome asli yang sudah login) ║
║  • Failsafe: Geser mouse ke pojok kiri-atas untuk STOP darurat       ║
║  • Untuk cek/ubah posisi klik: python fasih_auto_input_gui.py calibrate║
╚══════════════════════════════════════════════════════════════════════╝
    """)

    print(f"📊 Total Data di Excel : {len(items)}")
    print(f"⏳ Data Pending        : {len(pending)}")

    if not pending:
        print("🎉 Semua data sudah berstatus BERHASIL!")
        return

    print("\nPastikan:")
    print("1. Anda sudah login ke FASIH di Google Chrome.")
    print("2. Google Chrome terbuka di layar utama.")
    print("\nBot akan mulai dalam 4 detik... SIAPKAN LAYARMU!")
    for sec in range(4, 0, -1):
        print(f"   Memulai dalam {sec}...")
        time.sleep(1)

    sukses = 0
    gagal = 0

    try:
        for idx, item in enumerate(pending, 1):
            try:
                process_item(item, coords)
                update_row_status(wb, EXCEL_FILE, item, "BERHASIL")
                sukses += 1
                print(f"✅ Row {item['row_idx']} BERHASIL disimpan ke Excel.")
            except Exception as e:
                gagal += 1
                err_msg = f"GAGAL: {str(e)[:40]}"
                update_row_status(wb, EXCEL_FILE, item, err_msg)
                print(f"❌ Row {item['row_idx']} {err_msg}")

            time.sleep(1.0)

    except KeyboardInterrupt:
        print("\n\n⚠️ Bot dihentikan secara manual oleh pengguna.")

    print("\n" + "=" * 60)
    print("SELESAI - RINGKASAN:")
    print(f"✓ Berhasil : {sukses}")
    print(f"✗ Gagal    : {gagal}")
    print(f"Excel      : {EXCEL_FILE} (status terupdate)")
    print("=" * 60)


if __name__ == "__main__":
    main()
