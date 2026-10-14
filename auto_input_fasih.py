"""
FASIH Survey Automation - Input Data & Auto Submit
===================================================
Script otomasi Playwright untuk input data dari Excel (data_katA_kopi.xlsx) ke FASIH Web.

Alur Kerja:
  1. Browser terbuka (headless=False) agar user bisa login manual.
  2. Buka setiap link assignment di file Excel:
     - Cari dan klik tombol Edit (02_btn_edit.html)
     - Cari dan buka Sidepanel BLOK II (02_sidepanel_blok2.html)
     - Jika ada lebih dari satu usaha, pilih usaha sesuai 'Nama Usaha' di Excel
     - Cek Tahun Beroperasi (pertanyaan 25):
       * <= 2025: Isi 26.a (jika ada), 26.b (Biaya produksi), 26.d (Biaya operasional), 27.a (Nilai pendapatan),
                  serta tambahkan Catatan (Remark) di 27.a dari kolom REMARK FASIH
       * >= 2026: Isi 30.a (jika ada), 30.b (Biaya produksi), 30.d (Biaya operasional), 31.a (Nilai pendapatan),
                  serta tambahkan Catatan (Remark) di 31.a dari kolom REMARK FASIH
     - Klik tombol 'Kirim' di header (09_kirim.html)
     - Pada modal Kirim (10_modalkirim.html), klik 'Kirim'
     - Pada popup konfirmasi (11_konfirmasi.html), klik 'Konfirmasi'
     - Update status di Excel menjadi 'BERHASIL'
  3. Lanjut ke link berikutnya hingga selesai.

Cara Pakai:
  python auto_input_fasih.py
  (opsional: python auto_input_fasih.py --dry-run untuk simulasi tanpa klik submit akhir)
"""

import asyncio
import logging
import os
import sys
import re
import time
import subprocess
from datetime import datetime
from pathlib import Path
from playwright.async_api import (
    async_playwright,
    Page,
    BrowserContext,
    TimeoutError as PlaywrightTimeout,
)
import openpyxl
from playwright_stealth.stealth import Stealth

# ─────────────────────────── KONFIGURASI ───────────────────────────

BASE_URL       = "https://fasih-sm.bps.go.id"
DEFAULT_EXCEL  = "data_katA_kopi.xlsx"
AUTH_FILE      = "auth.json"
SLOW_MO        = 250        # ms jeda antar aksi
TIMEOUT        = 35_000     # ms timeout default
DRY_RUN        = False      # True = simulasi tanpa kirim data
ASSIGNMENT_ROW_TIMEOUT = 180 # Batas waktu maksimal per baris assignment (detik)

LOG_FILE     = "log.txt"
LAST_ROW_TXT = "log_terakhir.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger("fasih_input")


def update_last_row_txt(
    item: dict,
    status: str,
    note: str = "",
    current_idx: int = 0,
    total_items: int = 0,
    sukses: int = 0,
    gagal: int = 0,
):
    """
    Menulis / memperbarui file log_terakhir.txt dengan ringkasan khusus row terakhir yang diproses.
    """
    try:
        thn = item.get("tahun_beroperasi", "-")
        fokus = "26b, 26d, 27a (+ Catatan)" if (isinstance(thn, int) and thn <= 2025) else "30b, 30d, 31a (+ Catatan)"
        sisa = max(0, total_items - (sukses + gagal))

        catatan_re_run = ""
        if "SKIP" in status.upper() or "GAGAL" in status.upper():
            catatan_re_run = (
                "Keterangan Skip : Formulir TIDAK dikirim ke server. Baris ini tetap berstatus GAGAL\n"
                "                  di Excel sehingga otomatis tetap pending dan AKAN DIEKSEKUSI ULANG\n"
                "                  pada penjalanan script berikutnya.\n"
            )

        content = (
            f"==================================================\n"
            f"           STATUS PROSES ROW TERAKHIR\n"
            f"==================================================\n"
            f"Waktu Update     : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Baris Excel      : Baris {item.get('row_idx', '-')} di Excel (Data ke-{current_idx} dari {total_items})\n"
            f"Nama Usaha       : {item.get('nama_usaha', '-')}\n"
            f"Nama Assignment  : {item.get('nama_assignment', '-')}\n"
            f"Tahun Beroperasi : {thn} (Fokus: {fokus})\n"
            f"Status           : {status}\n"
            f"Alasan / Pesan   : {note if note else '-'}\n"
            f"{catatan_re_run}"
            f"Link Assignment  : {item.get('link_fasih', '-')}\n"
            f"--------------------------------------------------\n"
            f"RINGKASAN PROGRES:\n"
            f"  Total Data     : {total_items}\n"
            f"  ✓ Berhasil     : {sukses}\n"
            f"  ✗ Gagal/Skip   : {gagal}\n"
            f"  ⏳ Sisa Pending : {sisa}\n"
            f"==================================================\n"
        )
        with open(LAST_ROW_TXT, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        log.warning(f"Gagal menulis {LAST_ROW_TXT}: {e}")


# ─────────────────────────── HELPER FORMAT & PARSING ─────────────────

def get_screen_bounds() -> tuple[int, int]:
    """Mendapatkan resolusi layar aktif di macOS (default fallback 1680x1050)."""
    try:
        out = subprocess.check_output(
            ["osascript", "-e", 'tell application "Finder" to get bounds of window of desktop'],
            timeout=3
        ).decode().strip()
        parts = [int(p.strip()) for p in out.split(",")]
        w = parts[2] - parts[0]
        h = parts[3] - parts[1]
        if w > 800 and h > 600:
            return w, h
    except Exception:
        pass
    return 1680, 1050


def maximize_chrome_window(width: int, height: int):
    """Memaksimalkan posisi dan ukuran jendela Google Chrome di macOS."""
    try:
        script = f'tell application "Google Chrome" to set bounds of front window to {{0, 0, {width}, {height}}}'
        subprocess.run(["osascript", "-e", script], timeout=3, check=False)
    except Exception:
        pass


def clean_number(val) -> str:
    """Bersihkan nilai angka untuk input form."""
    if val is None:
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""
    try:
        # Hapus desimal nol jika float dari excel (misal 5900000.0 -> 5900000)
        f = float(val_str.replace(",", "."))
        if f.is_integer():
            return str(int(f))
        return str(f)
    except ValueError:
        return val_str


def parse_year(val) -> int:
    """Ambil tahun dari nilai kolom tahun beroperasi."""
    if val is None:
        return 2025
    val_str = str(val).strip()
    try:
        return int(float(val_str))
    except ValueError:
        # Jika teks, cari 4 digit angka
        m = re.search(r"\b(19\d\d|20\d\d)\b", val_str)
        if m:
            return int(m.group(1))
        return 2025


# ─────────────────────────── EXCEL HELPERS ───────────────────────────

def load_assignments(filepath: Path) -> tuple[openpyxl.Workbook, list[dict]]:
    """Membaca data assignment dari Excel."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    headers = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]

    # Pemetaan kolom dinamis berdasarkan teks header agar fleksibel jika kolom Excel berubah
    col_map = {}
    for idx, h in enumerate(headers, 1):
        if not h:
            continue
        h_str = str(h).strip().lower()
        if "id_assignment" in h_str:
            col_map["id_assignment"] = idx
        elif "link" in h_str:
            col_map["link_fasih"] = idx
        elif "nama assignment" in h_str:
            col_map["nama_assignment"] = idx
        elif "nama usaha" in h_str:
            col_map["nama_usaha"] = idx
        elif "kegiatan utama" in h_str:
            col_map["kegiatan_utama"] = idx
        elif "produk utama" in h_str:
            col_map["produk_utama"] = idx
        elif "tahun beroperasi" in h_str or "(25)" in h_str:
            col_map["tahun_beroperasi"] = idx
        elif "gaji" in h_str or "(26a)" in h_str or "(30a)" in h_str:
            col_map["adj_gaji"] = idx
        elif "b.prod" in h_str or "biaya prod" in h_str or "26b" in h_str or "30b" in h_str:
            col_map["adj_biaya_prod"] = idx
        elif "biaya operasional" in h_str or "biaya ops" in h_str or "26d" in h_str or "30d" in h_str:
            col_map["adj_biaya_ops"] = idx
        elif "pendapatan" in h_str or "27a" in h_str or "31a" in h_str:
            col_map["adj_pendapatan"] = idx
        elif "remark" in h_str:
            col_map["remark"] = idx
        elif "status" in h_str:
            col_map["status"] = idx
        elif "waktu" in h_str:
            col_map["waktu"] = idx

    # Kolom Status & Timestamp
    status_col = col_map.get("status")
    time_col = col_map.get("waktu")

    if not status_col:
        status_col = ws.max_column + 1
        ws.cell(row=1, column=status_col, value="Status Otomasi")
        col_map["status"] = status_col
    if not time_col:
        time_col = status_col + 1
        ws.cell(row=1, column=time_col, value="Waktu Eksekusi")
        col_map["waktu"] = time_col

    def get_val(r: int, key: str, default=""):
        c = col_map.get(key)
        if c is not None:
            v = ws.cell(r, c).value
            return "" if v is None else v
        return default

    items = []
    for r in range(2, ws.max_row + 1):
        link = get_val(r, "link_fasih")
        if not link or not str(link).strip().startswith("http"):
            continue

        item = {
            "row_idx": r,
            "id_assignment": str(get_val(r, "id_assignment")),
            "link_fasih": str(link).strip(),
            "nama_assignment": str(get_val(r, "nama_assignment")),
            "nama_usaha": str(get_val(r, "nama_usaha")),
            "kegiatan_utama": str(get_val(r, "kegiatan_utama")),
            "produk_utama": str(get_val(r, "produk_utama")),
            "tahun_beroperasi": parse_year(get_val(r, "tahun_beroperasi")),
            "adj_gaji": clean_number(get_val(r, "adj_gaji")),          # 26a atau 30a jika ada
            "adj_biaya_prod": clean_number(get_val(r, "adj_biaya_prod")),  # 26b atau 30b
            "adj_biaya_ops": clean_number(get_val(r, "adj_biaya_ops")),    # 26d atau 30d
            "adj_pendapatan": clean_number(get_val(r, "adj_pendapatan")),  # 27a atau 31a
            "remark": str(get_val(r, "remark")).strip(),
            "status": str(get_val(r, "status")).strip(),
            "status_col": status_col,
            "time_col": time_col,
        }
        items.append(item)

    return wb, items


def save_assignment_status(wb: openpyxl.Workbook, filepath: Path, item: dict, status: str, note: str = ""):
    """Update status per baris dan simpan ke file Excel."""
    ws = wb.active
    r = item["row_idx"]
    val = status if not note else f"{status}: {note}"
    ws.cell(row=r, column=item["status_col"], value=val)
    ws.cell(row=r, column=item["time_col"], value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.save(filepath)


# ─────────────────────────── FORM INTERACTION HELPERS ─────────────────

LOADING_SELECTORS = [
    "[class*='f:z-100']",
    "[class*='f:bg-black/20']",
    "[class*='bg-black/20'][class*='items-center']",
    "#FormGear-loader:not(.hidden)",
    "ngx-spinner:not([style*='display: none'])",
    "div.ngx-spinner-overlay",
    "div.swal2-loader",
    "div.swal2-loading",
    "[aria-busy='true']",
    ".spinner-border:visible",
]


async def wait_for_loading(page: Page, timeout_ms: int = 45_000, initial_delay_ms: int = 300) -> None:
    """
    Menunggu hingga seluruh indikator loading / spinner / overlay di FASIH & FormGear selesai.
    Jika masih loading, bot TIDAK AKAN melanjutkan ke tahap selanjutnya sampai loading benar-benar hilang.
    """
    if initial_delay_ms > 0:
        await page.wait_for_timeout(initial_delay_ms)

    start_time = time.time()
    max_wait = timeout_ms / 1000.0

    while (time.time() - start_time) < max_wait:
        is_loading = False
        detected_sel = None

        for sel in LOADING_SELECTORS:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible():
                    is_loading = True
                    detected_sel = sel
                    break
            except Exception:
                continue

        if not is_loading:
            return

        log.info(f"  ⏳ Masih loading ({detected_sel}), menunggu hingga selesai...")
        try:
            # Tunggu elemen loading tersebut menghilang dari layar
            await page.locator(detected_sel).first.wait_for(state="hidden", timeout=8_000)
        except Exception:
            pass

        await page.wait_for_timeout(500)

    log.warning("⚠️ Waktu tunggu loading melebihi batas, mencoba melanjutkan...")


async def click_edit_button(page: Page) -> bool:
    """Cari dan klik tombol edit (tabler-icon-edit) jika ada."""
    await wait_for_loading(page)
    log.info("Mencari tombol Edit (floating action button)...")
    edit_selectors = [
        "button:has(svg.tabler-icon-edit)",
        "svg.tabler-icon-edit",
        "div.f\\:fixed button:has(svg.tabler-icon-edit)",
        "button[data-state]:has(svg.tabler-icon-edit)",
    ]

    for sel in edit_selectors:
        try:
            btn = page.locator(sel).first
            if await btn.is_visible(timeout=2_500):
                await btn.scroll_into_view_if_needed()
                await btn.click()
                log.info("✅ Tombol Edit berhasil diklik.")
                await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=600)
                return True
        except Exception:
            continue

    log.info("ℹ️ Tombol Edit tidak ditemukan atau form sudah dalam mode edit.")
    return False


async def ensure_sidebar_open(page: Page) -> None:
    """Pastikan sidebar navigasi kuesioner terbuka (terutama jika layar sempit / mobile drawer)."""
    try:
        # Cek apakah item sidebar sudah terlihat
        sidebar_item = page.locator("div.fasih-form-sidebar, aside div[title*='BLOK']").first
        if await sidebar_item.is_visible():
            return

        # Coba klik tombol toggle sidebar (hamburger icon)
        toggle_btn = page.locator("button[aria-label='sidebar-toggle'], button:has(svg.feather-menu)").first
        if await toggle_btn.is_visible(timeout=2_000):
            log.info("Membuka drawer sidebar...")
            await toggle_btn.click()
            await page.wait_for_timeout(600)
    except Exception:
        pass


async def navigate_to_blok2(page: Page, nama_usaha: str = "") -> bool:
    """Cari sidepanel BLOK II dan buka usaha terkait."""
    await wait_for_loading(page)
    await ensure_sidebar_open(page)
    log.info("Mencari sidepanel 'SE2026 - L BLOK II'...")

    blok2_selectors = [
        "div[title*='BLOK II']",
        "div.fasih-form-sidebar div[title*='BLOK II']",
        "xpath=//div[contains(@class, 'fasih-form-sidebar')]//*[contains(@title, 'BLOK II') or contains(normalize-space(), 'BLOK II')]",
        "xpath=//*[contains(@title, 'BLOK II') or contains(normalize-space(), 'BLOK II')]",
        "#fasih-form-nav-next-button[title*='BLOK II']",
    ]

    blok2_clicked = False
    for sel in blok2_selectors:
        try:
            el = page.locator(sel).first
            if await el.is_visible(timeout=5_000):
                await el.scroll_into_view_if_needed()
                await el.click()
                blok2_clicked = True
                log.info("✅ Sidepanel BLOK II diklik.")
                await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)
                break
        except Exception:
            continue

    if not blok2_clicked:
        log.warning("⚠️ Sidepanel BLOK II belum terlihat langsung. Mencoba scroll container sidebar...")
        try:
            sidebar = page.locator("div.fasih-form-sidebar, aside").first
            await sidebar.evaluate("el => el.scrollBy(0, 300)")
            await page.wait_for_timeout(500)
            for sel in blok2_selectors:
                el = page.locator(sel).first
                if await el.is_visible(timeout=5_000):
                    await el.scroll_into_view_if_needed()
                    await el.click()
                    blok2_clicked = True
                    log.info("✅ Sidepanel BLOK II diklik setelah scroll.")
                    await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)
                    break
        except Exception:
            pass

    if not blok2_clicked:
        log.error("❌ Sidepanel BLOK II tidak ditemukan!")
        return False

    # Jika ada nama usaha spesifik, coba pilih sub-item usaha tersebut di bawah Blok II
    if nama_usaha:
        await wait_for_loading(page)
        clean_nama = nama_usaha.strip()
        log.info(f"Mencari sub-item usaha: '{clean_nama}'...")
        usaha_selectors = [
            f"div[title='{clean_nama}']",
            f"div[title*='{clean_nama}']",
            f"xpath=//div[contains(@class, 'fasih-form-sidebar')]//*[contains(@title, '{clean_nama}') or contains(normalize-space(), '{clean_nama}')]",
            f"div.fasih-form-sidebar:has-text('{clean_nama}')",
        ]
        for u_sel in usaha_selectors:
            try:
                u_el = page.locator(u_sel).first
                if await u_el.is_visible(timeout=6_000):
                    await u_el.scroll_into_view_if_needed()
                    await u_el.click()
                    log.info(f"✅ Sub-item usaha '{clean_nama}' dipilih.")
                    await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)
                    return True
            except Exception:
                continue

        log.warning(f"⚠️ Sub-item usaha '{clean_nama}' tidak ditemukan di sidebar. Melanjutkan ke isian...")

    return True


QUESTION_FIELD_MAP = {
    "26. a.": ["gaji"],
    "26. b.": ["biaya_produksi"],
    "26. d.": ["operasional"],
    "27. a.": ["nilai_pendapatan"],
    "30. a.": ["gaji", "30a"],
    "30. b.": ["biaya_produksi", "30b"],
    "30. d.": ["operasional", "30d"],
    "31. a.": ["nilai_pendapatan", "31a"],
}


async def fill_question_input(page: Page, q_prefix: str, value: str, timeout_ms: int = 25_000) -> bool:
    """
    Mengisi input untuk rincian pertanyaan tertentu (misal '26. b.', '26. d.', '27. a.').
    Jika input tidak ditemukan, fungsi return False agar baris bisa langsung di-skip.
    """
    if value is None or str(value).strip() == "":
        log.info(f"  [Skip] {q_prefix} (nilai kosong)")
        return True

    val_str = str(value).strip()
    await wait_for_loading(page, timeout_ms=20_000)
    log.info(f"  [Input] Mengisi {q_prefix} dengan '{val_str}'...")

    # Kumpulan selector kandidat: dari ID spesifik FormGear sampai XPath text node
    selectors = []

    # 1. Selector ID FormGear yang spesifik jika ada di mapping
    id_keys = QUESTION_FIELD_MAP.get(q_prefix, [])
    for k in id_keys:
        selectors.append(f"div[id*='{k}'] input:not([disabled])")

    # 2. Selector berbasis teks persis pada node teks (sangat presisi)
    selectors.extend([
        f"xpath=//*[text()[contains(., '{q_prefix}')]]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]//input",
        f"xpath=//div[contains(normalize-space(), '{q_prefix}')]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]//input",
        f"xpath=//*[contains(text(), '{q_prefix}')]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]//input",
    ])

    inp = None
    start_t = time.time()
    max_wait_sec = timeout_ms / 1000.0

    while (time.time() - start_t) < max_wait_sec:
        for sel in selectors:
            try:
                loc = page.locator(sel).first
                if await loc.is_visible(timeout=1_500):
                    inp = loc
                    break
            except Exception:
                continue

        if inp:
            break

        # Coba scroll pelan ke bawah jika elemen belum masuk render window
        try:
            await page.evaluate("window.scrollBy(0, 350)")
        except Exception:
            pass
        await page.wait_for_timeout(600)

    if not inp:
        log.error(f"  ❌ Timeout ({int(max_wait_sec)}s): Input untuk {q_prefix} tidak ditemukan!")
        return False

    try:
        await inp.scroll_into_view_if_needed()
        await page.wait_for_timeout(200)

        # Klik dan bersihkan nilai lama
        await inp.click()
        # Seleksi semua dan hapus
        await inp.press("Control+A" if sys.platform != "darwin" else "Meta+A")
        await inp.press("Backspace")
        await page.wait_for_timeout(100)

        # Isi nilai baru
        await inp.fill(val_str)
        # Trigger SolidJS event
        await inp.dispatch_event("input")
        await inp.dispatch_event("change")
        await inp.press("Tab")
        await wait_for_loading(page, timeout_ms=10_000, initial_delay_ms=300)
        log.info(f"  ✓ {q_prefix} berhasil diisi: {val_str}")
        return True
    except Exception as e:
        log.error(f"  ❌ Gagal mengisi {q_prefix}: {e}")
        return False


async def add_catatan_remark(page: Page, q_prefix: str, remark_text: str, timeout_ms: int = 25_000) -> bool:
    """
    Membuka modal Catatan (Remark) pada rincian tertentu, mengisi teks catatan, dan menyimpannya.
    Mendukung deteksi tombol visible baik pada tampilan layar lebar (desktop) maupun mobile.
    """
    if not remark_text:
        return True

    await wait_for_loading(page, timeout_ms=20_000)
    log.info(f"  [Catatan] Membuka modal Catatan untuk {q_prefix}...")

    # Cari container pertanyaan terlebih dahulu
    container_selectors = []
    id_keys = QUESTION_FIELD_MAP.get(q_prefix, [])
    for k in id_keys:
        container_selectors.append(f"div[id*='{k}']")
    container_selectors.extend([
        f"xpath=//*[text()[contains(., '{q_prefix}')]]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]",
        f"xpath=//div[contains(normalize-space(), '{q_prefix}')]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]",
    ])

    btn_catatan = None
    start_t = time.time()
    max_wait_sec = timeout_ms / 1000.0

    while (time.time() - start_t) < max_wait_sec:
        # Coba temukan tombol yang visible di dalam container pertanyaan
        for c_sel in container_selectors:
            try:
                c_loc = page.locator(c_sel).first
                if await c_loc.count() > 0:
                    catatan_btns = c_loc.locator("button[title='Catatan'], button[aria-label='Catatan']")
                    cnt = await catatan_btns.count()
                    for idx in range(cnt):
                        b = catatan_btns.nth(idx)
                        if await b.is_visible():
                            btn_catatan = b
                            break
            except Exception:
                continue
            if btn_catatan:
                break

        if btn_catatan:
            break

        # Fallback: cari tombol Catatan di sekitar teks q_prefix langsung
        try:
            nearby_btns = page.locator(
                f"xpath=//*[text()[contains(., '{q_prefix}')]]/ancestor::div[contains(@class, 'tw:rounded-lg')][1]//button[@title='Catatan' or @aria-label='Catatan']"
            )
            cnt = await nearby_btns.count()
            for idx in range(cnt):
                b = nearby_btns.nth(idx)
                if await b.is_visible():
                    btn_catatan = b
                    break
        except Exception:
            pass

        if btn_catatan:
            break

        # Scroll pelan ke bawah jika belum kelihatan
        try:
            await page.evaluate("window.scrollBy(0, 300)")
        except Exception:
            pass
        await page.wait_for_timeout(600)

    if not btn_catatan:
        log.warning(f"  ⚠️ Timeout ({int(max_wait_sec)}s): Tombol Catatan untuk {q_prefix} tidak ditemukan.")
        return False

    try:
        await btn_catatan.scroll_into_view_if_needed()
        await btn_catatan.click()
        await wait_for_loading(page, timeout_ms=10_000, initial_delay_ms=500)

        # Tunggu modal Catatan muncul
        modal_title = page.locator("#dialog-cl-13-title, div:has-text('Catatan'), [role='dialog']:has-text('Catatan')").first
        await modal_title.wait_for(state="visible", timeout=10_000)

        # Cek apakah catatan sudah pernah ditambahkan agar tidak dobel
        existing = await page.locator("div.tw\\:wrap-break-word, div.tw\\:whitespace-pre-wrap").all_text_contents()
        if any(remark_text in text for text in existing):
            log.info(f"  ℹ️ Catatan '{remark_text}' sudah ada di riwayat, tidak perlu simpan ulang.")
        else:
            # Ketik catatan baru
            ta = page.locator("textarea[placeholder*='catatan'], textarea[placeholder*='Catatan'], textarea").first
            await ta.wait_for(state="visible", timeout=8_000)
            await ta.fill(remark_text)
            await page.wait_for_timeout(300)

            # Klik tombol Simpan (icon paper plane dengan title Simpan)
            btn_simpan = page.locator("button[title='Simpan'], button:has-text('Simpan')").first
            await btn_simpan.wait_for(state="visible", timeout=6_000)
            await btn_simpan.click()
            log.info(f"  ✓ Catatan '{remark_text}' berhasil disimpan!")
            # Tunggu proses simpan catatan selesai di server
            await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)

        # Tutup modal
        btn_close = page.locator("button[aria-label='Dismiss'], button:has-text('Close')").first
        if await btn_close.is_visible():
            await btn_close.click()
        else:
            await page.keyboard.press("Escape")

        await wait_for_loading(page, timeout_ms=10_000, initial_delay_ms=400)
        return True

    except Exception as e:
        log.error(f"  ❌ Gagal menambahkan Catatan untuk {q_prefix}: {e}")
        try:
            await page.keyboard.press("Escape")
        except Exception:
            pass
        return False


async def check_and_resolve_galat(page: Page) -> bool:
    """
    Mengecek apakah terdapat Galat pada modal Kirim (10_modalkirim.html).
    Jika Galat > 0 (misalnya Anomali 'Wajib diisi'):
      1. Klik tombol Galat pada modal Kirim
      2. Tunggu modal rincian galat (11_modal_galatanomalo.html) muncul
      3. Klik tombol 'View' pada setiap item galat (misal 'Tampilkan Anomali Usaha dan Keluarga')
      4. Aktifkan toggle switch anomali (12_toogle.html)
      5. Buka kembali modal Kirim dari header agar galat menjadi 0
    Return True jika ada galat yang berhasil diselesaikan, False jika tidak ada galat.
    """
    await wait_for_loading(page, timeout_ms=15_000)

    galat_btn = page.locator("button:has(span:text-is('Galat'))").first
    if not await galat_btn.is_visible(timeout=4_000):
        return False

    # Periksa jumlah galat dari teks tombol atau badge
    galat_count = 0
    try:
        galat_text = await galat_btn.inner_text()
        m = re.search(r"Galat\s*(\d+)", galat_text, re.IGNORECASE)
        if m:
            galat_count = int(m.group(1))
        else:
            span_num = galat_btn.locator("span.tw\\:text-base").first
            if await span_num.is_visible():
                txt = (await span_num.inner_text()).strip()
                if txt.isdigit():
                    galat_count = int(txt)
    except Exception:
        pass

    if galat_count <= 0:
        log.info("ℹ️ Tidak ada Galat pada modal Kirim (Galat = 0).")
        return False

    log.warning(f"⚠️ Terdeteksi {galat_count} Galat pada modal Kirim! Membuka daftar galat untuk diperbaiki...")

    # Klik tombol Galat untuk membuka daftar rincian galat (11_modal_galatanomalo.html)
    await galat_btn.scroll_into_view_if_needed()
    await galat_btn.click()
    await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=600)

    # Tangani item galat
    resolved_any = False
    for attempt in range(max(1, galat_count)):
        view_selectors = [
            "button[title='View']",
            "button:has-text('View')",
            "div:has-text('Anomali') ~ div button",
            "button:has(svg path[d*='17.563'])",
            "div.tw\\:rounded-lg:has-text('Anomali') button",
        ]

        view_btn = None
        for v_sel in view_selectors:
            loc = page.locator(v_sel).first
            try:
                if await loc.is_visible(timeout=3_000):
                    view_btn = loc
                    break
            except Exception:
                continue

        if not view_btn:
            # Fallback klik card anomali langsung jika tombol view tidak terpisah
            anomali_card = page.locator("div[title*='Anomali'], div:has-text('Anomali')").first
            if await anomali_card.is_visible(timeout=2_000):
                view_btn = anomali_card
            else:
                log.warning("  ⚠️ Tombol View pada item galat tidak ditemukan.")
                break

        log.info(f"  [Galat] Mengklik tombol View pada item galat ke-{attempt + 1}...")
        await view_btn.scroll_into_view_if_needed()
        await view_btn.click()
        await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)

        # Cari switch toggle (12_toogle.html)
        switch_input = page.locator("input[role='switch']").first
        switch_control = page.locator(
            "div[id^='switch-'][id$='-control'], "
            "div.tw\\:cursor-pointer:has(div[id$='-thumb']), "
            "div[role='group']:has(input[role='switch']) div.tw\\:cursor-pointer, "
            "input[role='switch'] + div"
        ).first

        try:
            await page.wait_for_selector(
                "input[role='switch'], div[id^='switch-'][id$='-control'], div.tw\\:cursor-pointer:has(div[id$='-thumb'])",
                timeout=12_000,
            )
        except Exception:
            pass

        # Periksa apakah switch sudah aktif
        is_checked = False
        if await switch_input.count() > 0:
            aria = await switch_input.get_attribute("aria-checked")
            is_checked = (aria == "true")

        if not is_checked:
            log.info("  [Galat] Mengaktifkan toggle switch Anomali (Wajib diisi)...")
            if await switch_control.is_visible(timeout=3_000):
                await switch_control.scroll_into_view_if_needed()
                await switch_control.click()
            elif await switch_input.count() > 0:
                await switch_input.scroll_into_view_if_needed()
                await switch_input.click(force=True)

            await page.wait_for_timeout(300)
            if await switch_input.count() > 0:
                aria_after = await switch_input.get_attribute("aria-checked")
                if aria_after != "true":
                    await switch_input.click(force=True)

            await wait_for_loading(page, timeout_ms=15_000, initial_delay_ms=800)
            log.info("  ✓ Toggle switch Anomali berhasil diaktifkan!")
            resolved_any = True
        else:
            log.info("  ℹ️ Toggle switch Anomali sudah aktif.")

    # Jika berhasil mengaktifkan toggle, buka kembali modal Kirim dari header
    if resolved_any:
        log.info("Membuka kembali modal Kirim dari header setelah menyelesaikan Galat Anomali...")
        kirim_header_selectors = [
            "button:has(svg):has-text('Kirim')",
            "header button:has-text('Kirim')",
            "button.tw\\:h-9:has-text('Kirim')",
            "button:has-text('Kirim')",
        ]
        clicked_hdr = False
        for h_sel in kirim_header_selectors:
            try:
                h_btn = page.locator(h_sel).first
                if await h_btn.is_visible(timeout=3_000):
                    await h_btn.scroll_into_view_if_needed()
                    await h_btn.click()
                    clicked_hdr = True
                    log.info("✅ Tombol Kirim di header berhasil diklik kembali.")
                    break
            except Exception:
                continue

        if clicked_hdr:
            await wait_for_loading(page, timeout_ms=35_000, initial_delay_ms=1_000)

    return resolved_any


async def submit_assignment(page: Page, dry_run: bool = False) -> tuple[bool, str]:
    """
    Alur submit kuesioner:
      1. Tunggu loading selesai
      2. Klik tombol 'Kirim' di header (jika modal belum terbuka)
      3. Tunggu loading kalkulasi validasi form selesai
      4. Cek & selesaikan Galat jika ada (misal Galat Anomali 'Wajib diisi')
      5. Pada modal kirim, klik tombol 'Kirim' di dalam modal
      6. Pada dialog 'Konfirmasi Kirim', klik tombol 'Konfirmasi'
      7. Tunggu loading pengiriman form selesai
    """
    if dry_run:
        log.info("⏭️ DRY RUN aktif - tidak melakukan klik submit akhir.")
        return True, "DRY_RUN"

    # Pastikan tidak ada loading sebelum memulai submit
    await wait_for_loading(page, timeout_ms=30_000)
    log.info("🚀 Memulai proses Kirim formulir...")

    # 1. Cek apakah modal kirim sudah terbuka di layar
    modal_sudah_terbuka = False
    try:
        batal_btn = page.locator("button:has-text('Batal')").first
        if await batal_btn.is_visible(timeout=1_500):
            modal_sudah_terbuka = True
            log.info("ℹ️ Modal Kirim sudah terbuka di layar.")
    except Exception:
        pass

    # Jika modal belum terbuka, klik tombol Kirim di header (yang ada icon SVG pesawat)
    if not modal_sudah_terbuka:
        kirim_header_selectors = [
            "button:has(svg):has-text('Kirim')",
            "header button:has-text('Kirim')",
            "button.tw\\:h-9:has-text('Kirim')",
            "button:has-text('Kirim')",
        ]
        clicked_header = False
        for h_sel in kirim_header_selectors:
            try:
                h_btn = page.locator(h_sel).first
                if await h_btn.is_visible(timeout=2_000):
                    await h_btn.scroll_into_view_if_needed()
                    await h_btn.click()
                    clicked_header = True
                    log.info("✅ Tombol Kirim di header diklik.")
                    break
            except Exception:
                continue

        if not clicked_header:
            log.error("❌ Tombol Kirim header tidak ditemukan!")
            return False, "Kirim button header not found"

        # TUNGGU LOADING VALIDASI FORM (FormGear mengecek 305 jawaban)
        log.info("Menunggu validasi data & loading form kirim...")
        await wait_for_loading(page, timeout_ms=35_000, initial_delay_ms=800)

    # 1.5 CEK DAN SELESAIKAN GALAT ANOMALI (10_modalkirim.html -> 11_modal_galatanomalo.html -> 12_toogle.html)
    await check_and_resolve_galat(page)

    # 2. Klik tombol Kirim di dalam Modal Kirim (10_modalkirim.html)
    try:
        # Pastikan loading selesai sebelum berinteraksi dengan modal
        await wait_for_loading(page, timeout_ms=20_000)
        log.info("Mencari tombol Kirim di dalam modal...")

        modal_selectors = [
            "xpath=//button[normalize-space()='Batal']/parent::*//button[normalize-space()='Kirim']",
            "xpath=//button[normalize-space()='Batal']/..//button[normalize-space()='Kirim']",
            "xpath=//button[normalize-space()='Batal']/following-sibling::*//button[normalize-space()='Kirim']",
            "div.tw\\:sm\\:flex-row button:has-text('Kirim')",
            "div[role='dialog'] button:has-text('Kirim')",
        ]

        modal_kirim_btn = None
        for sel in modal_selectors:
            loc = page.locator(sel).last
            try:
                if await loc.is_visible(timeout=2_500):
                    modal_kirim_btn = loc
                    break
            except Exception:
                continue

        if not modal_kirim_btn:
            # Fallback tunggu selector utama
            fallback_loc = page.locator("xpath=//button[normalize-space()='Batal']/parent::*//button[normalize-space()='Kirim']").last
            await fallback_loc.wait_for(state="visible", timeout=10_000)
            modal_kirim_btn = fallback_loc

        await modal_kirim_btn.scroll_into_view_if_needed()
        try:
            await modal_kirim_btn.click(timeout=4_000)
        except Exception:
            # Force click jika backdrop atau animasi CSS menutupi pointer event
            log.info("Mencoba force click pada tombol Kirim modal...")
            await modal_kirim_btn.click(force=True)

        log.info("✅ Tombol Kirim di dalam modal berhasil diklik.")
        # TUNGGU LOADING SETELAH KLIK KIRIM MODAL
        await wait_for_loading(page, timeout_ms=25_000, initial_delay_ms=600)

    except Exception as e:
        log.error(f"❌ Gagal klik tombol Kirim di dalam modal: {e}")
        return False, f"Modal kirim error: {e}"

    # 3. Klik tombol Konfirmasi pada popup 'Konfirmasi Kirim' (11_konfirmasi.html)
    try:
        # Pastikan loading selesai sebelum klik konfirmasi
        await wait_for_loading(page, timeout_ms=20_000)
        log.info("Menunggu popup Konfirmasi Kirim...")

        confirm_selectors = [
            "xpath=//button[normalize-space()='Konfirmasi']",
            "button:has-text('Konfirmasi')",
            "div[role='dialog'] button:has-text('Konfirmasi')",
        ]

        confirm_btn = None
        for c_sel in confirm_selectors:
            loc = page.locator(c_sel).last
            try:
                await loc.wait_for(state="visible", timeout=15_000)
                confirm_btn = loc
                break
            except Exception:
                continue

        if not confirm_btn:
            # Cek apakah masih ada pesan galat atau peringatan di layar
            error_texts = await page.locator(".tw\\:text-destructive, [role='alert']").all_inner_texts()
            extra_info = " | ".join([t.strip() for t in error_texts if t.strip()])[:150]
            err_msg = f"Tombol Konfirmasi tidak muncul dalam batas waktu! {f'({extra_info})' if extra_info else ''}"
            raise RuntimeError(err_msg)

        await confirm_btn.scroll_into_view_if_needed()
        try:
            await confirm_btn.click(timeout=4_000)
        except Exception:
            log.info("Mencoba force click pada tombol Konfirmasi...")
            await confirm_btn.click(force=True)

        log.info("✅ Tombol Konfirmasi berhasil diklik.")

        # TUNGGU LOADING PENGIRIMAN ASSIGNMENT KE SERVER
        log.info("Menunggu proses pengiriman data ke server selesai...")
        await wait_for_loading(page, timeout_ms=45_000, initial_delay_ms=800)

        # Tunggu respon / modal tertutup
        try:
            await confirm_btn.wait_for(state="hidden", timeout=15_000)
        except Exception:
            pass

        # Cek jika ada notifikasi / alert penutup (SweetAlert OK / Tutup)
        try:
            ok_btn = page.locator("button.swal2-confirm, button:has-text('OK'), button:has-text('Tutup'), button:has-text('Selesai')").first
            if await ok_btn.is_visible(timeout=3_000):
                await ok_btn.click()
                log.info("✅ Notifikasi sukses ditutup.")
                await wait_for_loading(page, timeout_ms=10_000, initial_delay_ms=400)
        except Exception:
            pass

        await page.wait_for_timeout(2_000)
        log.info("🎉 Submit berhasil diselesaikan!")
        return True, "BERHASIL"

    except Exception as e:
        log.error(f"❌ Gagal klik tombol Konfirmasi: {e}")
        return False, f"Konfirmasi error: {e}"


# ─────────────────────────── PROCESS ONE ASSIGNMENT ──────────────────

async def process_single_assignment(page: Page, item: dict, dry_run: bool = False) -> tuple[bool, str]:
    """Menjalankan seluruh alur untuk satu baris assignment."""
    url = item["link_fasih"]
    thn = item["tahun_beroperasi"]
    usaha = item["nama_usaha"]

    log.info("=" * 60)
    log.info(f"Row {item['row_idx']}: {usaha} (Tahun: {thn})")
    log.info(f"URL: {url}")
    log.info("=" * 60)

    try:
        # 1. Buka URL assignment
        log.info("Membuka URL assignment...")
        await page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT)
        
        # Tunggu form kuesioner siap & loading awal selesai
        log.info("Menunggu form kuesioner siap (FormGear rendering)...")
        await wait_for_loading(page, timeout_ms=45_000, initial_delay_ms=1_000)
        try:
            await page.wait_for_selector(
                "button:has(svg.tabler-icon-edit), div.fasih-form-sidebar, div[title*='BLOK II']",
                state="visible",
                timeout=30_000,
            )
        except Exception:
            pass
        await wait_for_loading(page, timeout_ms=20_000, initial_delay_ms=500)

        # 2. Klik tombol Edit
        await click_edit_button(page)

        row_num = item["row_idx"]

        # 3. Buka Sidepanel BLOK II dan pilih usaha
        blok2_ok = await navigate_to_blok2(page, nama_usaha=usaha)
        if not blok2_ok:
            skip_msg = f"Sidepanel BLOK II / usaha '{usaha}' tidak ditemukan"
            log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
            return False, skip_msg

        # 4. Tentukan set pertanyaan berdasarkan Tahun Beroperasi
        if thn <= 2025:
            log.info(f"Tahun beroperasi {thn} <= 2025: Fokus ke rincian 26b, 26d, 27a (+ Catatan)")

            # 26. a. (jika ada nilai > 0)
            if item["adj_gaji"] and item["adj_gaji"] not in ("0", ""):
                if not await fill_question_input(page, "26. a.", item["adj_gaji"]):
                    skip_msg = "Input 26. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 26. b. Biaya produksi
            if item["adj_biaya_prod"]:
                if not await fill_question_input(page, "26. b.", item["adj_biaya_prod"]):
                    skip_msg = "Input 26. b. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 26. d. Biaya operasional
            if item["adj_biaya_ops"]:
                if not await fill_question_input(page, "26. d.", item["adj_biaya_ops"]):
                    skip_msg = "Input 26. d. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 27. a. Nilai pendapatan
            if item["adj_pendapatan"]:
                if not await fill_question_input(page, "27. a.", item["adj_pendapatan"]):
                    skip_msg = "Input 27. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # Catatan (Remark) di 27. a.
            if item["remark"]:
                if not await add_catatan_remark(page, "27. a.", item["remark"]):
                    skip_msg = "Catatan/Remark 27. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

        else:
            log.info(f"Tahun beroperasi {thn} >= 2026: Fokus ke rincian 30b, 30d, 31a (+ Catatan)")

            # 30. a. (jika ada nilai > 0)
            if item["adj_gaji"] and item["adj_gaji"] not in ("0", ""):
                if not await fill_question_input(page, "30. a.", item["adj_gaji"]):
                    skip_msg = "Input 30. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 30. b. Biaya produksi
            if item["adj_biaya_prod"]:
                if not await fill_question_input(page, "30. b.", item["adj_biaya_prod"]):
                    skip_msg = "Input 30. b. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 30. d. Biaya operasional
            if item["adj_biaya_ops"]:
                if not await fill_question_input(page, "30. d.", item["adj_biaya_ops"]):
                    skip_msg = "Input 30. d. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # 31. a. Nilai pendapatan
            if item["adj_pendapatan"]:
                if not await fill_question_input(page, "31. a.", item["adj_pendapatan"]):
                    skip_msg = "Input 31. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

            # Catatan (Remark) di 31. a.
            if item["remark"]:
                if not await add_catatan_remark(page, "31. a.", item["remark"]):
                    skip_msg = "Catatan/Remark 31. a. tidak ditemukan (Timeout 25s)"
                    log.warning(f"⚠️ [SKIP] Baris Excel {row_num} ({usaha}): {skip_msg}. Langsung di-skip & FORM TIDAK DIKIRIM!")
                    return False, skip_msg

        # 5. Kirim data (Hanya dikirim jika semua isian & catatan sukses terisi 100%)
        ok, msg = await submit_assignment(page, dry_run=dry_run)
        return ok, msg

    except PlaywrightTimeout as pe:
        msg = f"Timeout: {str(pe)[:100]}"
        log.error(f"❌ {msg}")
        try:
            await page.screenshot(path=f"error_row_{item['row_idx']}.png")
        except Exception:
            pass
        return False, msg

    except Exception as e:
        msg = f"Error: {str(e)[:100]}"
        log.error(f"❌ {msg}")
        try:
            await page.screenshot(path=f"error_row_{item['row_idx']}.png")
        except Exception:
            pass
        return False, msg


# ─────────────────────────── MAIN ORCHESTRATOR ───────────────────────

async def main():
    excel_path = Path(DEFAULT_EXCEL)
    if not excel_path.exists():
        log.error(f"File Excel '{excel_path}' tidak ditemukan!")
        return

    dry_run = "--dry-run" in sys.argv

    print(r"""
╔══════════════════════════════════════════════════════════════════════╗
║             FASIH AUTOMATION - INPUT DATA & AUTO SUBMIT              ║
║                                                                      ║
║  1. Browser Chrome akan terbuka otomatis                             ║
║  2. Silakan LOGIN MANUAL ke akun FASIH Anda di browser               ║
║  3. Setelah login & dashboard terbuka, tekan ENTER di terminal ini   ║
║  4. Bot akan otomatis menginput semua data Excel & kirim satu per satu║
╚══════════════════════════════════════════════════════════════════════╝
    """)

    wb, items = load_assignments(excel_path)
    pending_items = [it for it in items if it["status"] not in ("BERHASIL", "OK")]

    log.info(f"Total data di Excel : {len(items)}")
    log.info(f"Data pending        : {len(pending_items)}")
    log.info(f"Mode DRY_RUN        : {dry_run}")

    if not pending_items:
        log.info("🎉 Semua assignment di Excel sudah berstatus BERHASIL!")
        return

    # Inisialisasi storage_state jika ada di root atau di AutomationApproveFasih
    storage_state = None
    if Path(AUTH_FILE).exists():
        storage_state = AUTH_FILE
    elif Path("AutomationApproveFasih/auth.json").exists():
        storage_state = "AutomationApproveFasih/auth.json"

    chrome_proc = None
    caffeinate_proc = None
    profile_dir = Path.home() / "chrome_fasih_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    # 1. Cegah Mac masuk mode sleep selama bot berjalan
    if sys.platform == "darwin":
        try:
            caffeinate_proc = subprocess.Popen(["caffeinate", "-dims"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            log.info("☕ Caffeinate aktif: Mac dicegah masuk mode sleep selama bot berjalan.")
        except Exception as e:
            log.warning(f"Tidak dapat mengaktifkan caffeinate: {e}")

    try:
        # 2. Deteksi resolusi layar asli dan sesuaikan browser agar maximize
        screen_w, screen_h = get_screen_bounds()
        log.info(f"🖥️ Resolusi layar terdeteksi: {screen_w} x {screen_h}. Menyesuaikan browser agar maximize penuh...")

        # Buka aplikasi Google Chrome asli sistem persis seperti manusia membuka Chrome biasa
        log.info(f"Membuka Google Chrome di {BASE_URL}/ ...")
        chrome_cmd = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "--remote-debugging-port=9222",
            f"--user-data-dir={str(profile_dir)}",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized",
            f"--window-size={screen_w},{screen_h}",
            "--window-position=0,0",
            f"{BASE_URL}/"
        ]
        chrome_proc = subprocess.Popen(chrome_cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        await asyncio.sleep(4.0)

        async with async_playwright() as pw:
            log.info("Menghubungkan Playwright ke Google Chrome (CDP port 9222)...")
            browser = await pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
            context = browser.contexts[0]
            page = context.pages[0] if context.pages else await context.new_page()
            page.set_default_timeout(TIMEOUT)

            # Maksimalkan ukuran jendela Chrome dan viewport sesuai layar monitor
            maximize_chrome_window(screen_w, screen_h)
            try:
                dim = await page.evaluate("() => ({ w: window.screen.availWidth, h: window.screen.availHeight })")
                vw = dim.get("w") or screen_w
                vh = dim.get("h") or screen_h
                await page.set_viewport_size({"width": vw, "height": vh})
                maximize_chrome_window(vw, vh)
            except Exception:
                await page.set_viewport_size({"width": screen_w, "height": screen_h})

            # Cek apakah sudah login atau perlu login manual
            is_logged_in = False
            try:
                user_el = page.locator("span.user-name, .avatar-status-online, button[aria-label='sidebar-toggle'], div.fasih-form-sidebar").first
                if await user_el.is_visible(timeout=5_000):
                    is_logged_in = True
                    log.info("✅ Sudah dalam keadaan login!")
            except Exception:
                pass

            if not is_logged_in:
                print("\n" + "=" * 60)
                print("⏳ GOOGLE CHROME TERBUKA DI https://fasih-sm.bps.go.id/")
                print("Silakan LOGIN MANUAL seperti biasa.")
                print("Setelah login berhasil dan masuk ke FASIH,")
                print("TEKAN ENTER DI TERMINAL INI UNTUK MEMULAI OTOMASI INPUT...")
                print("=" * 60)
                input()

            # Simpan state login untuk penggunaan berikutnya
            try:
                await context.storage_state(path=AUTH_FILE)
                log.info(f"Session disimpan ke {AUTH_FILE}")
            except Exception:
                pass

            # Mulai perulangan setiap data
            sukses = 0
            gagal = 0

            for i, item in enumerate(pending_items, 1):
                log.info(f"\n[{i}/{len(pending_items)}] Memproses Baris {item['row_idx']} di Excel: {item['nama_usaha']}...")
                update_last_row_txt(
                    item,
                    "SEDANG DIPROSES...",
                    current_idx=i,
                    total_items=len(pending_items),
                    sukses=sukses,
                    gagal=gagal,
                )

                # Watchdog timer: jika proses pada 1 baris assignment stuck terlalu lama (> ASSIGNMENT_ROW_TIMEOUT detik),
                # script akan membatalkan baris tersebut dan langsung lanjut ke baris berikutnya.
                try:
                    ok, status_msg = await asyncio.wait_for(
                        process_single_assignment(page, item, dry_run=dry_run),
                        timeout=ASSIGNMENT_ROW_TIMEOUT
                    )
                except asyncio.TimeoutError:
                    ok = False
                    status_msg = f"Stuck/Timeout: Melebihi batas waktu maksimal ({ASSIGNMENT_ROW_TIMEOUT}s)"
                    log.error(
                        f"❌ [TIMEOUT / STUCK] Baris {item['row_idx']} ({item['nama_usaha']}) stuck lebih dari {ASSIGNMENT_ROW_TIMEOUT} detik! "
                        f"Langsung di-skip ke baris berikutnya..."
                    )

                if ok:
                    sukses += 1
                    save_assignment_status(wb, excel_path, item, "BERHASIL")
                    log.info(f"✅ Baris {item['row_idx']} di Excel SELESAI ({status_msg}). Status Excel di-update: BERHASIL.")
                    update_last_row_txt(
                        item,
                        "BERHASIL",
                        note=status_msg,
                        current_idx=i,
                        total_items=len(pending_items),
                        sukses=sukses,
                        gagal=gagal,
                    )
                else:
                    gagal += 1
                    save_assignment_status(wb, excel_path, item, "GAGAL", note=status_msg)
                    log.warning(
                        f"⚠️ Baris {item['row_idx']} di Excel ({item['nama_usaha']}) DI-SKIP!\n"
                        f"   Alasan: {status_msg}\n"
                        f"   Formulir TIDAK disubmit & status disimpan sebagai GAGAL (tetap pending) agar otomatis dieksekusi ulang saat script dijalankan kembali."
                    )
                    update_last_row_txt(
                        item,
                        "DI-SKIP (GAGAL)",
                        note=f"{status_msg} (Tetap pending di Excel untuk re-run)",
                        current_idx=i,
                        total_items=len(pending_items),
                        sukses=sukses,
                        gagal=gagal,
                    )

                await asyncio.sleep(1.0)

            print("\n" + "=" * 60)
            print("RINGKASAN EKSEKUSI:")
            print(f"Total diproses : {len(pending_items)}")
            print(f"✓ Berhasil      : {sukses}")
            print(f"✗ Gagal         : {gagal}")
            print(f"File Excel     : {excel_path} (status terupdate)")
            print(f"File Log       : {LOG_FILE}")
            print("=" * 60)

            input("\nTekan ENTER untuk selesai...")
            await browser.close()

    finally:
        if chrome_proc:
            try:
                chrome_proc.terminate()
            except Exception:
                pass
        if caffeinate_proc:
            try:
                caffeinate_proc.terminate()
                log.info("☕ Caffeinate dinonaktifkan: Mode daya PC kembali normal.")
            except Exception:
                pass


if __name__ == "__main__":
    asyncio.run(main())
