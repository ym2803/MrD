#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KS_UPDATE_INTEGRATED.py - Toko CGAR
Gabungan EXINOUT_CGAR.py (ambil data Google Sheets) + PrintAuto.py (otomasi EBarcode)

Alur baru (manual, tidak ada background/otomatis lagi):

  [UPDATE KS] (satu tombol, gabungan IN & OUT, deteksi otomatis)
      -> Ambil data spreadsheet SEKALI
      -> Kalau ada data IN  -> tulis EXCEL IN BOX.txt  & +PRINT IN LABEL.txt
                             -> baca EXCEL IN BOX.txt  -> jalankan otomasi EBarcode
                             -> centang "+PRINT IN LABEL"
      -> Kalau ada data OUT -> tulis EXCEL OUT BOX.txt & -PRINT OUT LABEL.txt
                             -> baca EXCEL OUT BOX.txt -> jalankan otomasi EBarcode
                             -> centang "-PRINT OUT LABEL"
      -> Kalau dua-duanya ada, dua-duanya diproses berurutan (IN dulu, baru OUT).

  [ULANGI PROSES EBARCODE] (untuk kasus lupa apakah EBarcode sudah dijalankan)
      -> TIDAK mengambil data baru dari Google Sheets.
      -> Langsung baca ulang EXCEL IN BOX.txt / EXCEL OUT BOX.txt yang sudah ada
         di disk, lalu ulangi dari langkah buka aplikasi EBarcode -> cetak.

Perubahan dibanding versi sebelumnya:
  - TIDAK ADA lagi watchdog folder 05 (data tidak diambil otomatis/background).
  - TIDAK ADA lagi jam operasional / auto-shutdown 09:00-22:00.
  - Data diambil dari Google Sheets HANYA saat tombol UPDATE KS diklik (manual).
  - Nama file & label tombol printing diubah dari "PRINT DISINI" menjadi
    "+PRINT IN LABEL" (proses IN) dan "-PRINT OUT LABEL" (proses OUT).
  - Menu "UPDATE KS IN" dan "UPDATE KS OUT" DIGABUNG menjadi satu tombol
    "UPDATE KS" yang otomatis mendeteksi ada data IN dan/atau OUT lalu
    menjalankan langkah yang sesuai (tidak perlu pilih IN/OUT secara manual).
  - Tombol baru "ULANGI PROSES EBARCODE" ditambahkan untuk mengulang HANYA
    tahap buka aplikasi EBarcode & cetak (tanpa ambil data GSheet lagi),
    untuk kasus user lupa sudah menjalankan otomasi EBarcode atau belum.
  - SEMUA klik/close/focus di otomasi EBarcode sekarang berbasis PostMessage
    (safe click), jadi tetap berfungsi walau PC dalam keadaan LOCKED.

PENGUATAN DATA PROCESSED (v2.2):
  - State disimpan di SQLite (<KODE>_state.db): commit inkremental, atomik, snapshot
    otomatis + pemulihan, riwayat sync. Format lama (JSON) dimigrasi otomatis.
  - Spreadsheet dibaca hemat memori (hanya kolom B..BQ, baris kosong dibuang).
  - Verifikasi pembacaan kedua: perubahan yang tidak stabil (mis. sedang cut-paste)
    ditunda ke Sync berikutnya, bukan dianggap IN/OUT.

PERBAIKAN STABILITAS SYNC (v2.1):
  - Sumber data dibaca UTUH (tanpa batas baris). Dulu RANGE "B6:BQ2000" membuat
    data di bawah baris 2000 tidak pernah terbaca -> IN tidak jalan & baris yang
    bergeser lewat 2000 dianggap OUT palsu.
  - Parser CSV memakai modul csv (aman untuk sel multi-baris), fetch dengan retry,
    deteksi halaman HTML/login, dan validasi data.
  - State ditulis atomik + backup, per-toko, dan DI-COMMIT per batch HANYA setelah
    cetak EBarcode sukses (kalau gagal, data otomatis muncul lagi di Sync berikut).
  - Pengaman OUT massal (mencegah ribuan OUT palsu kalau sheet gagal terbaca penuh).
  - Data banyak dipecah per batch (BATCH_SIZE) saat otomasi EBarcode.
  - Virtual Desktop tidak bocor saat proses gagal; screenshot debug dibatasi.

Requires: pip install pywinauto pyvda
"""

import os
import re
import sys
import json
import time
import ctypes
import logging
import subprocess
import csv
import io
import socket
import shutil
import sqlite3
import http.client
import urllib.error
import urllib.request
import urllib.parse
import threading
import webbrowser
from pathlib import Path
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import scrolledtext, messagebox, ttk, simpledialog
    HAS_TK = True
except ImportError:
    HAS_TK = False

try:
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys
    HAS_PYWINAUTO = True
except ImportError:
    HAS_PYWINAUTO = False

try:
    from pyvda import VirtualDesktop, AppView
    PYVDA_OK = True
except ImportError:
    PYVDA_OK = False


# --- DPI Awareness ---------------------------------------------------------
def _set_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_set_dpi_awareness()


# ============================================================
# KONFIGURASI
# ============================================================

def _app_dir() -> Path:
    """
    Folder tempat aplikasi 'berdiri':
    - Jika dijalankan sebagai .exe hasil PyInstaller (--onefile / --onedir),
      __file__ bisa menunjuk ke folder ekstraksi sementara (sys._MEIPASS)
      yang dihapus setiap kali exe ditutup. Karena itu kita pakai
      sys.executable, yang selalu menunjuk ke lokasi exe sebenarnya.
    - Jika dijalankan sebagai script .py biasa, tetap pakai folder script.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _app_dir()

CONFIG_FILE = APP_DIR / "config_ks.json"

DEFAULT_STORES = [
    {"code": "CGAR", "spreadsheet_id": "1ii5ka5ABGmNrp-15rm0T-Xb2-aEkiSb3J2Py8gfBn78", "gid": "628762981"},
]

# Kode permanen untuk membuka menu Pengaturan (klik ikon gerigi).
# Tidak disimpan di config_ks.json supaya tidak bisa diubah/dilihat lewat file itu.
SETTINGS_PASSCODE = "yayangmrdiycgar"

# Nomor WA untuk minta kode akses Pengaturan (format internasional untuk wa.me)
SETTINGS_WA_NUMBER = "6282115303443"

DEFAULT_PATHS = {
    "dir_scan":   r"D:\Scanning",
    "dir_proses": r"D:\QasDev\QubeV10\BackEnd\Data\Seuic\06",
    "exe_path":   r"D:\QasDev\QubeV10\BackEnd\EBarcode.exe",
}


def _normalize_stores(raw) -> list:
    stores, seen = [], set()
    for item in raw or []:
        code = str(item.get("code", "")).strip()
        sid  = str(item.get("spreadsheet_id", "")).strip()
        gid  = str(item.get("gid", "")).strip()
        if code and sid and gid and code not in seen:
            stores.append({"code": code, "spreadsheet_id": sid, "gid": gid})
            seen.add(code)
    return stores


def load_config() -> dict:
    stores = [s.copy() for s in DEFAULT_STORES]
    active = stores[0]["code"]
    paths = DEFAULT_PATHS.copy()
    store_selected = False
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)

            store_selected = bool(saved.get("store_selected", False))

            if "stores" in saved:
                loaded = _normalize_stores(saved.get("stores"))
                if loaded:
                    stores = loaded
            elif saved.get("spreadsheet_id"):
                migrated = _normalize_stores([{
                    "code":           saved.get("store_code", "CGAR"),
                    "spreadsheet_id": saved.get("spreadsheet_id"),
                    "gid":            saved.get("gid"),
                }])
                if migrated:
                    stores = migrated

            codes = [s["code"] for s in stores]
            active_candidate = str(saved.get("active_store", "")).strip()
            active = active_candidate if active_candidate in codes else codes[0]

            paths = DEFAULT_PATHS.copy()
            saved_paths = saved.get("paths") or {}
            for k in DEFAULT_PATHS:
                if saved_paths.get(k):
                    paths[k] = str(saved_paths[k]).strip()
    except Exception as e:
        print(f"Gagal load config ({CONFIG_FILE}): {e}")
        paths = DEFAULT_PATHS.copy()
    return {"stores": stores, "active_store": active, "paths": paths, "store_selected": store_selected}


def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"Gagal simpan config ({CONFIG_FILE}): {e}")


_CONFIG           = load_config()
STORES            = _CONFIG["stores"]
ACTIVE_STORE_CODE = _CONFIG["active_store"]
_PATHS            = _CONFIG["paths"]
STORE_SELECTED    = _CONFIG["store_selected"]


def _find_store(code: str):
    for s in STORES:
        if s["code"] == code:
            return s
    return None


def _resolve_active_globals():
    global STORE_CODE, SPREADSHEET_ID, GID
    store = _find_store(ACTIVE_STORE_CODE) or (STORES[0] if STORES else DEFAULT_STORES[0])
    STORE_CODE     = store["code"]
    SPREADSHEET_ID = store["spreadsheet_id"]
    GID            = store["gid"]


def _persist_all():
    save_config({
        "stores": STORES,
        "active_store": ACTIVE_STORE_CODE,
        "paths": _PATHS,
        "store_selected": STORE_SELECTED,
    })


def mark_store_selected():
    global STORE_SELECTED
    STORE_SELECTED = True
    _persist_all()


def set_active_store(code: str):
    global ACTIVE_STORE_CODE
    if _find_store(code):
        ACTIVE_STORE_CODE = code
        _resolve_active_globals()
        _persist_all()


def upsert_store(code: str, spreadsheet_id: str, gid: str,
                  original_code: str = None, activate: bool = False):
    global ACTIVE_STORE_CODE

    code, spreadsheet_id, gid = code.strip(), spreadsheet_id.strip(), gid.strip()
    if not code or not spreadsheet_id or not gid:
        raise ValueError("Semua field wajib diisi.")

    existing_idx = None
    if original_code:
        for i, s in enumerate(STORES):
            if s["code"] == original_code:
                existing_idx = i
                break

    for i, s in enumerate(STORES):
        if s["code"] == code and i != existing_idx:
            raise ValueError(f"Kode toko '{code}' sudah dipakai.")

    new_entry = {"code": code, "spreadsheet_id": spreadsheet_id, "gid": gid}
    if existing_idx is not None:
        was_active = STORES[existing_idx]["code"] == ACTIVE_STORE_CODE
        STORES[existing_idx] = new_entry
        if was_active:
            ACTIVE_STORE_CODE = code
    else:
        STORES.append(new_entry)
        if activate or len(STORES) == 1:
            ACTIVE_STORE_CODE = code

    _resolve_active_globals()
    _persist_all()


def delete_store(code: str):
    global ACTIVE_STORE_CODE
    if len(STORES) <= 1:
        raise ValueError("Minimal harus ada 1 toko.")

    idx = next((i for i, s in enumerate(STORES) if s["code"] == code), None)
    if idx is None:
        return
    STORES.pop(idx)
    if ACTIVE_STORE_CODE == code:
        ACTIVE_STORE_CODE = STORES[0]["code"]

    _resolve_active_globals()
    _persist_all()


def apply_path_config(dir_scan: str, dir_proses: str, exe_path: str):
    global DIR_SCAN, DIR_PROSES, EXE_PATH, FILE1, FILE2, FILE3, FILE4, FILE_STATE

    dir_scan, dir_proses, exe_path = dir_scan.strip(), dir_proses.strip(), exe_path.strip()
    if not dir_scan or not dir_proses or not exe_path:
        raise ValueError("Semua field folder/aplikasi wajib diisi.")

    _PATHS["dir_scan"]   = dir_scan
    _PATHS["dir_proses"] = dir_proses
    _PATHS["exe_path"]   = exe_path
    _persist_all()

    DIR_SCAN   = dir_scan
    DIR_PROSES = dir_proses
    EXE_PATH   = exe_path

    FILE1 = os.path.join(DIR_PROSES, "EXCEL IN BOX.txt")
    FILE2 = os.path.join(DIR_SCAN,   "+PRINT IN LABEL.txt")
    FILE3 = os.path.join(DIR_PROSES, "EXCEL OUT BOX.txt")
    FILE4 = os.path.join(DIR_SCAN,   "-PRINT OUT LABEL.txt")
    FILE_STATE = os.path.join(DIR_PROSES, "CGAR_processed.txt")


# --- Google Sheets (sumber data EXINOUT) - diambil dari toko yang aktif ---
STORE_CODE     = None
SPREADSHEET_ID = None
GID            = None
_resolve_active_globals()
# Data dibaca dari SELURUH sheet (tanpa batas baris), lalu dipotong di sisi lokal.
# (Versi lama memakai RANGE "B6:BQ2000" => hanya ~1995 baris terbaca.)
SHEET_FIRST_ROW = 6   # baris data pertama (1-based) -> dulu "B6"
SHEET_FIRST_COL = 1   # kolom B (A=0, B=1) -> cols[0] pada semua indeks di bawah = kolom B


B_COL_INDEX  = 0
C_COL_INDEX  = 1
AH_COL_INDEX = 32
AI_COL_INDEX = 33

# --- Kolom Data Master (untuk KS RECON), offset dari kolom B (sama sheet/range) ---
W_COL_INDEX  = 21   # Kolom W  - SKU (Master)
X_COL_INDEX  = 22   # Kolom X  - Rack (Master)
AC_COL_INDEX = 27   # Kolom AC - No Ctn (Master)
AE_COL_INDEX = 29   # Kolom AE - Deskripsi (Master)

# --- Kolom Sync Recon (status per-baris), offset dari kolom B (sama sheet/range) ---
BM_COL_INDEX = 63   # Kolom BM - SKU
BN_COL_INDEX = 64   # Kolom BN - No Ctn / Box
BQ_COL_INDEX = 67   # Kolom BQ - Status Rekonsiliasi

STATUS_BELUM_IN  = "SKU Belum Excel In"
STATUS_BELUM_OUT = "SKU Belum Excel Out"

DIR_SCAN   = _PATHS["dir_scan"]
DIR_PROSES = _PATHS["dir_proses"]

FILE1 = os.path.join(DIR_PROSES, "EXCEL IN BOX.txt")
FILE2 = os.path.join(DIR_SCAN,   "+PRINT IN LABEL.txt")
FILE3 = os.path.join(DIR_PROSES, "EXCEL OUT BOX.txt")
FILE4 = os.path.join(DIR_SCAN,   "-PRINT OUT LABEL.txt")
FILE_STATE = os.path.join(DIR_PROSES, "CGAR_processed.txt")

EXE_PATH       = _PATHS["exe_path"]
TIMEOUT        = 25
LOG_FILE       = APP_DIR / "log.txt"

# --- Parameter stabilitas ---
BATCH_SIZE        = 150      # maks baris per putaran otomasi EBarcode
MAX_DEAD_KEYS     = 20000    # riwayat kunci OUT yang disimpan (hanya untuk label RE-IN)
FETCH_RETRIES     = 3
FETCH_TIMEOUT     = 90       # detik
GUARD_MIN_ACTIVE  = 30       # pengaman OUT massal aktif kalau data aktif >= nilai ini
GUARD_OUT_RATIO   = 0.5      # batalkan sync bila > 50% data aktif akan jadi OUT sekaligus
FORCE_SYNC_FLAG   = APP_DIR / "force_sync.flag"   # buat file kosong ini utk melewati pengaman 1x
MAX_DEBUG_SHOTS   = 30       # jumlah screenshot debug yang disimpan
FIRST_RUN_BASELINE = True    # True: sync pertama kali hanya MENCATAT data yang ada (tidak dicetak)
STOP_CLOSES_EBARCODE = True  # True: saat STOP, jendela EBarcode yang dibuka otomasi ikut ditutup
LEGACY_STATE_NAME = "CGAR_processed.txt"
STABLE_READ_DELAY = 4.0      # detik; saat ada perubahan, sheet dibaca ulang & hanya yang stabil diproses (0 = matikan)
KEEP_SYNC_HISTORY = 500      # jumlah riwayat sync yang disimpan
SHEET_URL_TEMPLATE = "https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"
FAMILY_MATCH_SKU = True    # True: proteksi 'family' harus cocok base + SKU yang sama

LABEL_SEARCH_IN  = "PRINT IN LABEL"
LABEL_SEARCH_OUT = "PRINT OUT LABEL"

BROWN_BOX_LABEL_NAME = "BROWN BOX"
BROWN_BOX_SKU        = "3101270"

BROWN_BOX_DROPDOWN_ITEMS = [
    "Default Format",
    "STOCK LABEL",
    "PRICE TAG",
    "COLLECTING LABEL",
    "GONDOLA",
    "BROWN BOX",
    "Default Landscape",
    "STOCK LABELLandscape",
]
BROWN_BOX_ITEM_INDEX = BROWN_BOX_DROPDOWN_ITEMS.index("BROWN BOX")

# Lookup teks (UPPER) -> index, dipakai untuk verifikasi & koreksi posisi
# combo "Select Label" (lihat open_brown_box_print_form).
BROWN_BOX_ITEMS_BY_TEXT = {
    item.upper(): idx for idx, item in enumerate(BROWN_BOX_DROPDOWN_ITEMS)
}

GONDOLA_LABEL_NAME = "GONDOLA"
GONDOLA_SKU        = BROWN_BOX_SKU
GONDOLA_ITEM_INDEX = BROWN_BOX_DROPDOWN_ITEMS.index("GONDOLA")


# ============================================================
# LOGGING
# ============================================================

def _write_crash_fallback(text: str):
    """Pencatat cadangan paling dasar (tidak pakai modul logging sama sekali),
    dipakai kalau setup logging sendiri gagal atau ada error fatal sebelum
    logging siap. Ditulis ke %TEMP% supaya selalu ada lokasi yang writable
    walau folder aplikasi (mis. di Program Files) tidak bisa ditulis."""
    try:
        import tempfile
        fallback_path = Path(tempfile.gettempdir()) / "exio_crash.log"
        with open(fallback_path, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {text}\n")
    except Exception:
        pass  # kalau ini pun gagal, sudah tidak ada lagi yang bisa dilakukan


_log_handlers = []

try:
    _log_handlers.append(logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8"))
except Exception as e:
    _write_crash_fallback(f"Gagal buat FileHandler di {LOG_FILE}: {e}")
    try:
        # Fallback: tulis log ke %TEMP% kalau folder aplikasi tidak writable
        # (mis. exe ditaruh di Program Files tanpa hak admin).
        import tempfile
        fallback_log = Path(tempfile.gettempdir()) / "exio_log.txt"
        _log_handlers.append(logging.FileHandler(fallback_log, mode="w", encoding="utf-8"))
    except Exception as e2:
        _write_crash_fallback(f"Gagal juga buat FileHandler fallback di TEMP: {e2}")

# StreamHandler ke console HANYA ditambahkan kalau memang ada console
# (sys.stdout tidak None). Exe hasil build --noconsole/--windowed punya
# sys.stdout = None, dan StreamHandler(None) bisa memicu error saat log
# pertama ditulis -- di titik ini errornya terjadi SEBELUM main() jalan,
# jadi aplikasi terlihat 'tidak membuka sama sekali' tanpa jejak apa pun.
if sys.stdout is not None:
    try:
        _log_handlers.append(logging.StreamHandler(sys.stdout))
    except Exception as e:
        _write_crash_fallback(f"Gagal buat StreamHandler: {e}")

try:
    logging.basicConfig(
        level=logging.INFO,
        format="[%(asctime)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=_log_handlers,
        force=True,
    )
except Exception as e:
    _write_crash_fallback(f"Gagal basicConfig logging: {e}")

log = logging.getLogger("KS_UPDATE")


def _global_excepthook(exc_type, exc_value, exc_tb):
    """Jaring pengaman terakhir: menangkap error apa pun yang tidak
    tertangani di mana pun dalam program (termasuk sebelum GUI sempat
    terbuka) dan selalu mencatatnya, baik lewat logging maupun lewat
    penulis cadangan yang tidak bergantung pada logging."""
    import traceback
    text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    try:
        log.error(f"UNCAUGHT EXCEPTION:\n{text}")
    except Exception:
        pass
    _write_crash_fallback(f"UNCAUGHT EXCEPTION:\n{text}")


sys.excepthook = _global_excepthook


def ui_log(msg, level=logging.INFO):
    log.log(level, msg, extra={"show_ui": True})


# ============================================================
# KONTROL JEDA / STOP (tombol di aplikasi + hotkey global)
# ============================================================
_real_sleep = time.sleep
_real_time = time.time


class StopRequested(BaseException):
    """Sengaja turunan BaseException (bukan Exception) supaya tidak ikut
    tertelan oleh blok 'except Exception' di dalam alur otomasi EBarcode."""


class RunControl:
    def __init__(self):
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._run = threading.Event()
        self._run.set()                 # set = berjalan, clear = dijeda
        self._active = False
        self._cleanup = False
        self._paused = False
        self._pause_started = 0.0
        self._paused_total = 0.0

    # --- siklus hidup satu proses ---
    def begin(self):
        with self._lock:
            self._stop.clear()
            self._run.set()
            self._active, self._cleanup = True, False
            self._paused, self._paused_total = False, 0.0

    def end(self):
        with self._lock:
            self._active = False
            self._run.set()
            self._paused = False

    # --- status ---
    @property
    def active(self): return self._active
    @property
    def paused(self): return self._paused
    @property
    def stop_requested(self): return self._stop.is_set()

    # --- perintah dari UI / hotkey ---
    def set_paused(self, value: bool):
        with self._lock:
            if value and not self._paused:
                self._paused, self._pause_started = True, _real_time()
                self._run.clear()
            elif not value and self._paused:
                self._paused_total += _real_time() - self._pause_started
                self._paused = False
                self._run.set()

    def request_stop(self):
        self._stop.set()
        self.set_paused(False)          # lepas jeda supaya alur bisa keluar

    def paused_total(self) -> float:
        with self._lock:
            extra = (_real_time() - self._pause_started) if self._paused else 0.0
            return self._paused_total + extra

    # --- dipanggil dari alur otomasi (thread kerja) ---
    def checkpoint(self):
        if not self._active or self._cleanup:
            return
        if threading.current_thread() is threading.main_thread():
            return
        while True:
            if self._stop.is_set():
                raise StopRequested()
            if self._run.is_set():
                return
            _real_sleep(0.1)

    def cleanup(self):
        """Dipakai saat membersihkan (mis. menutup EBarcode) setelah STOP,
        supaya langkah pembersihan itu sendiri tidak ikut dihentikan."""
        ctl = self
        class _Ctx:
            def __enter__(s): ctl._cleanup = True
            def __exit__(s, *a): ctl._cleanup = False
        return _Ctx()


control = RunControl()


def ctl_sleep(seconds: float):
    """Pengganti time.sleep: dipecah kecil supaya JEDA/STOP langsung terasa."""
    control.checkpoint()
    remaining = max(0.0, float(seconds))
    while remaining > 0:
        step = min(0.1, remaining)
        _real_sleep(step)
        remaining -= step
        control.checkpoint()


def ctl_time() -> float:
    """Pengganti time.time untuk hitung timeout: waktu selama JEDA tidak ikut
    dihitung, jadi jeda lama tidak membuat langkah otomasi 'timeout'."""
    return _real_time() - control.paused_total()


def start_global_hotkeys(on_pause, on_stop) -> bool:
    """Hotkey global (berfungsi walau jendela aplikasi tidak sedang fokus):
       Ctrl+Alt+P = Jeda/Lanjut     Ctrl+Alt+S = Stop
    (Windows tidak mengirim hotkey saat PC terkunci/lock screen.)"""
    if os.name != "nt":
        return False

    def _loop():
        try:
            import ctypes.wintypes as wt
            user32 = ctypes.windll.user32
            mods = 0x0002 | 0x0001 | 0x4000      # CTRL | ALT | NOREPEAT
            if not user32.RegisterHotKey(None, 1, mods, ord("P")):
                log.warning("Hotkey Ctrl+Alt+P gagal didaftarkan (dipakai aplikasi lain?) - pakai tombol.")
            if not user32.RegisterHotKey(None, 2, mods, ord("S")):
                log.warning("Hotkey Ctrl+Alt+S gagal didaftarkan (dipakai aplikasi lain?) - pakai tombol.")
            msg = wt.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                if msg.message == 0x0312:         # WM_HOTKEY
                    if msg.wParam == 1:
                        on_pause()
                    elif msg.wParam == 2:
                        on_stop()
        except Exception as e:
            log.warning(f"Hotkey global tidak aktif: {e}")

    threading.Thread(target=_loop, daemon=True, name="exio-hotkeys").start()
    return True


_automation_desktop: object = None
_user_desktop: object = None


# ============================================================
# KLIK & CLOSE BERBASIS MESSAGE — tetap jalan walau PC LOCKED
# click_input()/send_keys() bawaan pywinauto memakai SendInput/mouse_event
# (simulasi input hardware) yang DIBLOKIR Windows saat workstation locked.
# PostMessage tidak butuh 'interactive input desktop', jadi tetap berfungsi.
# ============================================================
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP   = 0x0202
WM_SYSCOMMAND  = 0x0112
SC_CLOSE       = 0xF060


def _make_lparam(x: int, y: int) -> int:
    return (int(y) << 16) | (int(x) & 0xFFFF)


def click_message(hwnd, x: int, y: int, delay: float = 0.015):
    """Klik di koordinat (x, y) relatif ke client-area hwnd, via PostMessage.
    Tetap berfungsi walau PC dalam keadaan locked/terkunci."""
    try:
        lparam = _make_lparam(x, y)
        user32 = ctypes.windll.user32
        user32.PostMessageW(hwnd, WM_LBUTTONDOWN, 1, lparam)
        ctl_sleep(delay)
        user32.PostMessageW(hwnd, WM_LBUTTONUP, 0, lparam)
        return True
    except Exception as e:
        log.warning(f"click_message gagal pada hwnd={hwnd}: {e}")
        return False


def click_control_message(control, delay: float = 0.03):
    """Klik di tengah sebuah control pywinauto tanpa perlu koordinat manual."""
    try:
        rect = control.rectangle()
        cx = (rect.right - rect.left) // 2
        cy = (rect.bottom - rect.top) // 2
        return click_message(control.handle, cx, cy, delay)
    except Exception as e:
        log.warning(f"click_control_message gagal: {e}")
        return False


def close_window_message(hwnd):
    """Tutup window via WM_SYSCOMMAND/SC_CLOSE — pengganti send_keys('%{F4}')
    yang gagal saat PC locked karena butuh SendInput."""
    try:
        ctypes.windll.user32.PostMessageW(hwnd, WM_SYSCOMMAND, SC_CLOSE, 0)
        return True
    except Exception as e:
        log.warning(f"close_window_message gagal pada hwnd={hwnd}: {e}")
        return False


def safe_set_focus(control):
    """set_focus() bisa gagal/exception saat PC locked (butuh SetForegroundWindow).
    Bungkus supaya tidak menghentikan proses — klik berbasis message tidak
    butuh focus untuk berhasil."""
    try:
        control.set_focus()
        return True
    except Exception as e:
        log.warning(f"set_focus dilewati (kemungkinan PC locked): {e}")
        return False


def find_deepest_control_at_point(root_control, screen_x, screen_y):
    """Cari control (dengan HWND asli) yang PALING SPESIFIK menutupi titik
    (screen_x, screen_y). Dipakai untuk klik owner-drawn seperti label
    'Excel In' yang tidak selalu tergambar langsung di atas form."""
    best = root_control
    best_area = None
    try:
        candidates = [root_control] + list(root_control.descendants())
    except Exception:
        candidates = [root_control]

    for ctrl in candidates:
        try:
            r = ctrl.rectangle()
            if r.left <= screen_x <= r.right and r.top <= screen_y <= r.bottom:
                area = (r.right - r.left) * (r.bottom - r.top)
                if area <= 0:
                    continue
                if best_area is None or area < best_area:
                    best = ctrl
                    best_area = area
        except Exception:
            pass
    return best


def click_screen_point(root_control, screen_x, screen_y, delay: float = 0.03):
    """Klik di titik (screen_x, screen_y) dengan menargetkan HWND paling
    spesifik di titik itu (bukan selalu window paling luar)."""
    target = find_deepest_control_at_point(root_control, screen_x, screen_y)
    try:
        t_rect = target.rectangle()
        rel_x = screen_x - t_rect.left
        rel_y = screen_y - t_rect.top
        log.info(f"  [click_screen_point] target={target.class_name()} "
                 f"handle={target.handle} rect={t_rect} rel=({rel_x},{rel_y})")
        return click_message(target.handle, rel_x, rel_y, delay)
    except Exception as e:
        log.warning(f"click_screen_point gagal: {e}")
        return False


def save_debug_screenshot(hwnd, tag: str):
    """Screenshot window via PrintWindow API (tetap jalan walau window ada
    di virtual desktop tersembunyi/PC locked). Dipakai untuk kalibrasi
    koordinat klik saat langkah berbasis-koordinat gagal."""
    try:
        import win32gui, win32ui
        from ctypes import windll
        from PIL import Image

        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        w, h = right - left, bottom - top
        if w <= 0 or h <= 0:
            return None

        hwnd_dc = win32gui.GetWindowDC(hwnd)
        mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
        save_dc = mfc_dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(mfc_dc, w, h)
        save_dc.SelectObject(bmp)

        windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

        bmp_info = bmp.GetInfo()
        bmp_bits = bmp.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]),
            bmp_bits, "raw", "BGRX", 0, 1
        )

        out_dir = APP_DIR / "debug_screenshots"
        out_dir.mkdir(exist_ok=True)
        try:
            shots = sorted(out_dir.glob("*.png"), key=lambda p: p.stat().st_mtime)
            for old in shots[:-MAX_DEBUG_SHOTS]:
                old.unlink()
        except Exception:
            pass
        out_path = out_dir / f"{tag}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        img.save(out_path)

        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

        log.info(f"  [DEBUG] Screenshot disimpan: {out_path}")
        return out_path
    except Exception as e:
        log.warning(f"  [DEBUG] Gagal ambil screenshot: {e}")
        return None


def click_excel_in_with_retry(frm_print, nav_rect, grid_rect, max_wait_per_try=2.0):
    """'Excel In' adalah label/link owner-drawn tanpa HWND sendiri — bisa saja
    digambar di atas panel/groupbox lain (bukan langsung di form), jadi klik
    harus diarahkan ke HWND paling spesifik di titik tsb (lihat
    click_screen_point), bukan cuma ditembak ke frm_print.handle. Dicoba
    beberapa titik kandidat sampai popup/dialog berikutnya terdeteksi."""
    nav_h = nav_rect.bottom - nav_rect.top
    excel_abs_y = nav_rect.top + nav_h // 2

    candidates = []
    if grid_rect:
        gap = grid_rect.right - nav_rect.right
        for ratio in (0.55, 0.35, 0.7, 0.45, 0.85):
            candidates.append(nav_rect.right + int(gap * ratio))
    for offset in (30, 15, 45, 60, 80):
        x = nav_rect.right + offset
        if x not in candidates:
            candidates.append(x)

    for i, excel_abs_x in enumerate(candidates):
        log.info(f"  [Percobaan {i+1}/{len(candidates)}] Klik Excel In di titik "
                 f"abs=({excel_abs_x},{excel_abs_y})")
        click_screen_point(frm_print, excel_abs_x, excel_abs_y)

        deadline = ctl_time() + max_wait_per_try
        while ctl_time() < deadline:
            if Desktop(backend="win32").windows(class_name="TMessageForm"):
                log.info("  OK: Popup konfirmasi muncul, titik klik ini benar")
                return True
            if Desktop(backend="win32").windows(class_name="TdlgStoreSHELF"):
                log.info("  OK: TdlgStoreSHELF langsung muncul, titik klik ini benar")
                return True
            ctl_sleep(0.05)

    log.error("  Semua titik kandidat gagal memicu 'Excel In'")
    save_debug_screenshot(frm_print.handle, "excel_in_gagal")
    return False


# ============================================================
# NOTICE POPUP
# ============================================================
def show_notice(success: bool, message: str, auto_close_sec: int = 10, stopped: bool = False):
    if not HAS_TK:
        print(f"\n{'[BERHASIL]' if success else '[GAGAL]'} {message}")
        return

    root = tk.Toplevel() if tk._default_root else tk.Tk()
    root.title(f"{STORE_CODE} \u2013 " + ("Berhasil" if success else ("Dihentikan" if stopped else "Gagal")))
    root.resizable(False, False)
    root.attributes("-topmost", True)

    BG     = "#FFFFFF"
    ACCENT = "#1976D2" if success else ("#F57C00" if stopped else "#D32F2F")
    FG_MSG = "#16324F"
    FG_CNT = "#5A7897"

    root.configure(bg=BG)

    tk.Frame(root, bg=ACCENT, height=6).pack(fill="x")

    frame_top = tk.Frame(root, bg=BG, padx=30, pady=20)
    frame_top.pack(fill="x")
    icon_char = "\u2714" if success else ("\u25A0" if stopped else "\u2718")
    judul     = "Proses Selesai" if success else ("Proses Dihentikan" if stopped else "Proses Gagal")
    tk.Label(frame_top, text=icon_char, font=("Segoe UI", 28, "bold"),
             fg=ACCENT, bg=BG).pack(side="left", padx=(0, 14))
    tk.Label(frame_top, text=judul, font=("Segoe UI", 14, "bold"),
             fg="#0D2A4A", bg=BG).pack(side="left", anchor="w")

    frame_msg = tk.Frame(root, bg=BG, padx=30)
    frame_msg.pack(fill="x")
    tk.Label(frame_msg, text=message, wraplength=380, justify="left",
             font=("Segoe UI", 10), fg=FG_MSG, bg=BG).pack(anchor="w")

    frame_bot = tk.Frame(root, bg=BG, padx=30, pady=18)
    frame_bot.pack(fill="x")
    lbl_count = tk.Label(frame_bot, text="", font=("Segoe UI", 9),
                         fg=FG_CNT, bg=BG)
    lbl_count.pack(side="left")
    tk.Button(frame_bot, text="  OK  ", font=("Segoe UI", 10, "bold"),
              bg=ACCENT, fg="#ffffff", relief="flat",
              activebackground=ACCENT, cursor="hand2",
              command=root.destroy).pack(side="right")

    tk.Frame(root, bg=ACCENT, height=4).pack(fill="x", side="bottom")

    remaining = [auto_close_sec]

    def tick():
        if remaining[0] <= 0:
            root.destroy()
            return
        lbl_count.config(text=f"Menutup otomatis dalam {remaining[0]} detik...")
        remaining[0] -= 1
        root.after(1000, tick)

    tick()

    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")


# ============================================================
# BAGIAN 1: SINKRONISASI DATA (dari EXINOUT_CGAR.py)
# ============================================================

def parse_csv_line(line):
    cols, current, in_quotes = [], [], False
    for char in line:
        if char == '"':
            in_quotes = not in_quotes
        elif char == ',' and not in_quotes:
            cols.append(''.join(current).strip().strip('"').strip())
            current = []
        else:
            current.append(char)
    cols.append(''.join(current).strip().strip('"').strip())
    return [c.strip() for c in cols]


def _atomic_write_text(filepath, text: str):
    """Tulis file secara atomik (tulis ke file sementara lalu os.replace) supaya
    tidak pernah menghasilkan file setengah jadi/rusak kalau PC mati atau proses
    terhenti. Nama file sementara sengaja TIDAK memuat teks label, supaya tidak
    ikut terbaca di daftar EBarcode."""
    folder = os.path.dirname(filepath) or "."
    os.makedirs(folder, exist_ok=True)
    tmp = os.path.join(folder, f"~exio_{os.getpid()}.tmp")
    try:
        with open(tmp, "w", encoding="utf-8", newline="") as f:
            f.write(text)
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
        for attempt in range(6):
            try:
                os.replace(tmp, filepath)
                return
            except PermissionError:
                if attempt == 5:
                    raise
                ctl_sleep(0.3 * (attempt + 1))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass


def write_txt(filepath, lines):
    _atomic_write_text(filepath, '\r\n'.join(lines))


def b_base(b_val: str) -> str:
    parts = b_val.split(' ', 1)
    kode  = parts[0]
    sisa  = parts[1] if len(parts) == 2 else None

    kode_bersih = re.sub(r'[A-Za-z]+$', '', kode)

    if not kode_bersih:
        return b_val

    if sisa is not None:
        return f"{kode_bersih} {sisa}"
    else:
        return kode_bersih


def make_key(b, c):
    return f"{b}|{c}"


def state_file_path() -> str:
    """Path state FORMAT LAMA (JSON/teks). Sekarang hanya dipakai untuk migrasi."""
    return os.path.join(DIR_PROSES, f"{STORE_CODE}_processed.txt")


def state_db_path() -> str:
    """State sekarang berupa database SQLite per toko."""
    return os.path.join(DIR_PROSES, f"{STORE_CODE}_state.db")


class SyncState:
    """Status kunci (No Ctn|SKU) yang sudah diproses, disimpan di SQLite.

    Kenapa SQLite (bukan satu file JSON):
      - Commit INKREMENTAL: hanya baris yang berubah yang ditulis, bukan menulis
        ulang seluruh file tiap batch (aman & cepat walau data ratusan ribu).
      - Transaksi atomik, tahan mati listrik / proses terhenti di tengah tulis.
      - Snapshot otomatis (2 generasi) + pemulihan otomatis kalau file rusak.
      - Riwayat sync tersimpan (untuk menelusuri kejadian).
    Perubahan HANYA disimpan lewat commit_in/commit_out, dipanggil setelah cetak
    EBarcode sukses -> kalau cetak gagal, data otomatis muncul lagi di Sync berikutnya.
    """

    def __init__(self):
        self.active = set()    # kunci yang sudah di-IN dan belum OUT
        self.dead = {}         # riwayat kunci OUT (berurutan, dibatasi MAX_DEAD_KEYS)
        self.is_new = True     # True = belum pernah ada state (pemakaian pertama)
        self._snapped = False

    # ---------------- koneksi & skema ----------------
    @staticmethod
    def _connect(path=None):
        conn = sqlite3.connect(path or state_db_path(), timeout=30)
        conn.execute("PRAGMA journal_mode=DELETE")   # tanpa file -wal/-shm yang tertinggal
        conn.execute("PRAGMA synchronous=FULL")
        return conn

    @staticmethod
    def _init_schema(conn):
        conn.execute("CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT)")
        conn.execute("CREATE TABLE IF NOT EXISTS keys("
                     "k TEXT PRIMARY KEY, status TEXT NOT NULL, updated REAL NOT NULL)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_keys_status ON keys(status, updated)")
        conn.execute("CREATE TABLE IF NOT EXISTS sync_log("
                     "id INTEGER PRIMARY KEY AUTOINCREMENT, ts REAL, result TEXT, note TEXT, "
                     "rows INTEGER, active_before INTEGER, n_in INTEGER, n_out INTEGER, "
                     "n_move INTEGER, n_stay INTEGER, n_dup INTEGER, n_invalid INTEGER, "
                     "n_unstable INTEGER)")

    @classmethod
    def _read(cls, path):
        """Baca & validasi satu file database. Return (meta, rows)."""
        conn = cls._connect(path)
        try:
            cls._init_schema(conn)
            chk = conn.execute("PRAGMA quick_check").fetchone()
            if not chk or chk[0] != "ok":
                raise sqlite3.DatabaseError(f"quick_check gagal: {chk}")
            meta = dict(conn.execute("SELECT k, v FROM meta").fetchall())
            rows = conn.execute("SELECT k, status FROM keys ORDER BY updated, k").fetchall()
        finally:
            conn.close()
        return meta, rows

    # ---------------- migrasi dari format lama ----------------
    @classmethod
    def _read_legacy(cls):
        """Baca state format lama (JSON / teks). Return (active, dead, used_files) atau None."""
        path = state_file_path()
        bak = path + ".bak"
        legacy = os.path.join(DIR_PROSES, LEGACY_STATE_NAME)

        if os.path.exists(path) or os.path.exists(bak):
            candidates = [(path, "utama"), (bak, "backup")]
        elif (os.path.normcase(legacy) != os.path.normcase(path) and os.path.exists(legacy)):
            if len(STORES) == 1:
                candidates = [(legacy, "lama")]
            else:
                raise Exception(
                    f"Ditemukan state lama '{LEGACY_STATE_NAME}' tetapi belum ada state untuk toko "
                    f"'{STORE_CODE}'.\nJika toko ini SUDAH dipakai sebelum update, salin file itu "
                    f"menjadi '{os.path.basename(path)}' (folder yang sama) lalu Sync lagi.\n"
                    f"Jika toko ini BARU, pindahkan file lama itu ke tempat lain dulu.")
        else:
            return None

        for cand, label in candidates:
            if not os.path.exists(cand):
                continue
            try:
                with open(cand, "r", encoding="utf-8") as f:
                    raw = f.read()
                if not raw.strip():
                    raise ValueError("file kosong")
                if raw.lstrip().startswith("{"):
                    data = json.loads(raw)
                    sid, gid = data.get("spreadsheet_id"), data.get("gid")
                    if (sid and sid != SPREADSHEET_ID) or (gid and str(gid) != str(GID)):
                        raise LookupError(
                            f"State ini milik spreadsheet lain (id={sid}, gid={gid}). "
                            f"Toko '{STORE_CODE}' sekarang memakai id={SPREADSHEET_ID}, gid={GID}. "
                            f"Jika spreadsheet memang diganti, hapus file: {cand}")
                    active, dead = list(data.get("active", [])), list(data.get("dead", []))
                else:
                    active = [l.strip() for l in raw.splitlines() if "|" in l]
                    dead = []
                log.warning(f"State format lama ({label}) dimigrasi ke SQLite: "
                            f"{len(active)} aktif, {len(dead)} riwayat OUT.")
                return active, dead, [p for p in (path, bak, cand) if os.path.exists(p)]
            except LookupError as e:
                raise Exception(str(e))
            except Exception as e:
                log.warning(f"State lama ({label}) tidak bisa dibaca: {e}")

        raise Exception(
            "File state lama rusak dan backup tidak tersedia/tidak valid:\n"
            f"{path}\nSync dihentikan supaya tidak terjadi IN/OUT massal palsu. "
            "Jika ingin mulai dari nol, hapus file tersebut (data yang ada saat itu akan "
            "dicatat sebagai baseline, TIDAK dicetak).")

    @classmethod
    def _create_from(cls, db, legacy):
        """Buat database baru (opsional diisi dari state lama) dalam satu transaksi."""
        conn = cls._connect(db)
        try:
            with conn:
                cls._init_schema(conn)
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('store', ?)", (STORE_CODE,))
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('spreadsheet_id', ?)", (SPREADSHEET_ID,))
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('gid', ?)", (str(GID),))
                if legacy is not None:
                    active, dead, _files = legacy
                    now = _real_time()
                    conn.executemany("INSERT OR REPLACE INTO keys(k,status,updated) VALUES (?,'A',?)",
                                     ((k, now) for k in active))
                    conn.executemany("INSERT OR REPLACE INTO keys(k,status,updated) VALUES (?,'D',?)",
                                     ((k, now - 1 + i * 1e-6) for i, k in enumerate(dead)))
                    conn.execute("INSERT OR REPLACE INTO meta VALUES ('initialized', '1')")
        except BaseException:
            conn.close()
            try:
                os.remove(db)            # jangan tinggalkan database setengah jadi
            except Exception:
                pass
            raise
        conn.close()

    # ---------------- load + pemulihan ----------------
    @classmethod
    def load(cls):
        os.makedirs(DIR_PROSES, exist_ok=True)
        db = state_db_path()

        if not os.path.exists(db):
            legacy = cls._read_legacy()          # None = pemakaian pertama
            cls._create_from(db, legacy)
            if legacy is not None:
                for p in legacy[2]:              # arsipkan file lama (jangan dihapus)
                    try:
                        os.replace(p, p + ".migrated")
                    except Exception as e:
                        log.warning(f"Gagal mengarsipkan {p}: {e}")

        try:
            meta, rows = cls._read(db)
        except Exception as e:
            meta, rows = cls._recover(db, e)

        sid, gid = meta.get("spreadsheet_id"), meta.get("gid")
        if (sid and sid != SPREADSHEET_ID) or (gid and str(gid) != str(GID)):
            raise Exception(
                f"State ini milik spreadsheet lain (id={sid}, gid={gid}). "
                f"Toko '{STORE_CODE}' sekarang memakai id={SPREADSHEET_ID}, gid={GID}. "
                f"Jika spreadsheet memang diganti, hapus file: {db}")

        st = cls()
        for k, status in rows:
            if status == "A":
                st.active.add(k)
            else:
                st.dead[k] = None
        st.is_new = (meta.get("initialized") != "1")
        return st

    @classmethod
    def _recover(cls, db, err):
        log.warning(f"Database state bermasalah ({err}). Mencoba memulihkan dari snapshot...")
        for snap in (db + ".snap1", db + ".snap2"):
            if not os.path.exists(snap):
                continue
            try:
                meta, rows = cls._read(snap)
            except Exception as e:
                log.warning(f"  Snapshot {os.path.basename(snap)} tidak valid: {e}")
                continue
            try:
                os.replace(db, f"{db}.corrupt-{datetime.now():%Y%m%d-%H%M%S}")
            except Exception:
                pass
            shutil.copy2(snap, db)
            log.warning(f"  Database dipulihkan dari {os.path.basename(snap)} "
                        f"(perubahan sesudah snapshot itu bisa muncul lagi sebagai IN/OUT).")
            return meta, rows
        raise Exception(
            "File state (SQLite) rusak dan snapshot tidak tersedia/tidak valid:\n"
            f"{db}\nSync dihentikan supaya tidak terjadi IN/OUT massal palsu. "
            "Jika ingin mulai dari nol, hapus file tersebut (data yang ada saat itu akan "
            "dicatat sebagai baseline, TIDAK dicetak).")

    # ---------------- snapshot ----------------
    def _ensure_snapshot(self):
        """Salin kondisi SEBELUM perubahan pertama pada proses ini (maks 2 generasi)."""
        if self._snapped:
            return
        self._snapped = True
        db = state_db_path()
        tmp = db + ".snap.tmp"
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
            src, dst = self._connect(), sqlite3.connect(tmp)
            try:
                src.backup(dst)
            finally:
                dst.close()
                src.close()
            s1, s2 = db + ".snap1", db + ".snap2"
            if os.path.exists(s1):
                os.replace(s1, s2)
            os.replace(tmp, s1)
        except Exception as e:
            log.warning(f"Gagal membuat snapshot state: {e}")

    # ---------------- perubahan data ----------------
    def baseline(self, keys):
        """Catat semua kunci yang ada sebagai 'sudah diproses' (tanpa cetak)."""
        keys = list(keys)
        now = _real_time()
        conn = self._connect()
        try:
            with conn:
                conn.execute("DELETE FROM keys")
                conn.executemany("INSERT OR REPLACE INTO keys(k,status,updated) VALUES (?,'A',?)",
                                 ((k, now) for k in keys))
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('initialized', '1')")
        finally:
            conn.close()
        self.active, self.dead, self.is_new = set(keys), {}, False

    def commit_in(self, keys):
        keys = list(keys)
        if not keys:
            return
        self._ensure_snapshot()
        now = _real_time()
        conn = self._connect()
        try:
            with conn:
                conn.executemany("INSERT OR REPLACE INTO keys(k,status,updated) VALUES (?,'A',?)",
                                 ((k, now) for k in keys))
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('initialized', '1')")
        finally:
            conn.close()
        for k in keys:
            self.active.add(k)
            self.dead.pop(k, None)
        self.is_new = False

    def commit_out(self, keys):
        keys = list(keys)
        if not keys:
            return
        self._ensure_snapshot()
        now = _real_time()
        conn = self._connect()
        try:
            with conn:
                conn.executemany("INSERT OR REPLACE INTO keys(k,status,updated) VALUES (?,'D',?)",
                                 ((k, now + i * 1e-6) for i, k in enumerate(keys)))
                conn.execute(
                    "DELETE FROM keys WHERE status='D' AND k IN ("
                    "SELECT k FROM keys WHERE status='D' ORDER BY updated DESC, k DESC "
                    "LIMIT -1 OFFSET ?)", (MAX_DEAD_KEYS,))
                conn.execute("INSERT OR REPLACE INTO meta VALUES ('initialized', '1')")
        finally:
            conn.close()
        for k in keys:
            self.active.discard(k)
            self.dead.pop(k, None)
            self.dead[k] = None
        while len(self.dead) > MAX_DEAD_KEYS:
            self.dead.pop(next(iter(self.dead)))
        self.is_new = False

    # ---------------- riwayat ----------------
    def log_sync(self, result, note="", **c):
        """Catat ringkasan satu kali sync (tidak pernah melempar error)."""
        try:
            conn = self._connect()
            try:
                with conn:
                    conn.execute(
                        "INSERT INTO sync_log(ts,result,note,rows,active_before,n_in,n_out,n_move,"
                        "n_stay,n_dup,n_invalid,n_unstable) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                        (_real_time(), result, (note or "")[:500], c.get("rows", 0),
                         c.get("active_before", 0), c.get("n_in", 0), c.get("n_out", 0),
                         c.get("n_move", 0), c.get("n_stay", 0), c.get("n_dup", 0),
                         c.get("n_invalid", 0), c.get("n_unstable", 0)))
                    conn.execute("DELETE FROM sync_log WHERE id <= "
                                 "(SELECT MAX(id) FROM sync_log) - ?", (KEEP_SYNC_HISTORY,))
            finally:
                conn.close()
        except Exception as e:
            log.warning(f"Gagal mencatat riwayat sync: {e}")

    def history(self, n=10):
        try:
            conn = self._connect()
            try:
                rows = conn.execute(
                    "SELECT ts,result,note,rows,active_before,n_in,n_out,n_move,n_stay,n_dup,"
                    "n_invalid,n_unstable FROM sync_log ORDER BY id DESC LIMIT ?", (n,)).fetchall()
            finally:
                conn.close()
        except Exception:
            return []
        out = []
        for ts, res, note, rr, ab, i, o, mv, st, du, inv, un in rows:
            line = (f"{datetime.fromtimestamp(ts):%Y-%m-%d %H:%M:%S} | {res:<8} | baca {rr} | "
                    f"aktif {ab} | IN {i} OUT {o} MOVE {mv} ACTIVE {st} DUP {du} INVALID {inv} "
                    f"TUNDA {un}")
            if note:
                line += f" | {note}"
            out.append(line)
        return out


def write_status_report(stats: dict, report_path: str, history=None):
    order = [
        ("IN",        "DATA BARU (IN)"),
        ("RE_IN",     "DATA MUNCUL KEMBALI (RE-IN)"),
        ("MOVE",      "PINDAH LOKASI (MOVE) - No Ctn lama, SKU tetap hidup"),
        ("ACTIVE",    "TETAP AKTIF (ACTIVE) - keluarga SKU/No Ctn masih hidup"),
        ("OUT",       "DATA KELUAR (OUT) - sudah tidak ada relasi aktif"),
        ("DUPLICATE", "DATA DUPLIKAT (DUPLICATE) - No Ctn+SKU sama > 1 kali"),
        ("INVALID",   "DATA TIDAK VALID (INVALID)"),
    ]
    lines = []
    lines.append("=" * 70)
    lines.append(f"LAPORAN STATUS SINKRONISASI - {STORE_CODE}")
    lines.append(f"Waktu: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 70)
    for key_name, title in order:
        items = stats.get(key_name, [])
        lines.append("")
        lines.append(f"[{title}]  (total: {len(items)})")
        if items:
            for it in sorted(items):
                if "|" in it:
                    ctn, sku = it.rsplit("|", 1)
                    lines.append(f"    No Ctn: {ctn:<12} SKU: {sku}")
                else:
                    lines.append(f"    {it}")
        else:
            lines.append("    (tidak ada)")
    if history:
        lines.append("")
        lines.append("[RIWAYAT SYNC TERAKHIR]")
        for h in history:
            lines.append("    " + h)
    lines.append("")
    lines.append("=" * 70)

    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    with open(report_path, 'w', encoding='utf-8', newline='') as f:
        f.write('\n'.join(lines))


class SheetFetchError(Exception):
    pass


def _parse_sheet_csv(raw: bytes, keep_cols: int) -> list:
    """Parse CSV secara bertahap (tanpa membuat string raksasa) dan HANYA menyimpan
    kolom yang dipakai (B..BQ). Baris kosong dibuang. Hemat memori untuk sheet besar."""
    csv.field_size_limit(10_000_000)
    reader = csv.reader(io.TextIOWrapper(io.BytesIO(raw), encoding="utf-8-sig",
                                         errors="replace", newline=""))
    out = []
    for n, r in enumerate(reader, start=1):
        if n < SHEET_FIRST_ROW:
            continue
        cols = [c.strip() for c in r[SHEET_FIRST_COL:SHEET_FIRST_COL + keep_cols]]
        if any(cols):
            out.append(cols)
    if not out:
        raise SheetFetchError("Spreadsheet terbaca tetapi tidak ada isi data.")
    return out


def fetch_sheet_rows() -> list:
    """Ambil SELURUH sheet toko aktif sebagai list baris (tiap baris = list kolom,
    kolom pertama = kolom B, hanya sampai kolom BQ, baris kosong dibuang).
    - Tanpa batas baris (dulu B6:BQ2000).
    - Retry dengan jeda bertahap untuk error jaringan / 429 / 5xx / unduhan terpotong.
    - Menolak halaman HTML (login/error) supaya tidak dibaca sebagai data kosong.
    - Parsing memakai modul csv (benar untuk sel multi-baris & tanda kutip) dan
      hanya menyimpan kolom yang dipakai (jauh lebih hemat RAM)."""
    csv_url = SHEET_URL_TEMPLATE.format(sid=SPREADSHEET_ID, gid=GID)
    keep_cols = BQ_COL_INDEX + 1
    data, last_err = None, None

    for attempt in range(1, FETCH_RETRIES + 1):
        try:
            req = urllib.request.Request(csv_url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=FETCH_TIMEOUT) as resp:
                raw = resp.read()          # IncompleteRead bila unduhan terpotong
            head = raw[:600].decode("utf-8-sig", errors="replace").lstrip().lower()
            if head.startswith("<!doctype") or "<html" in head:
                raise SheetFetchError(
                    "Google mengembalikan halaman HTML, bukan CSV. Pastikan spreadsheet "
                    "di-share (Anyone with the link) dan Spreadsheet ID/GID benar.")
            if not raw.strip():
                raise SheetFetchError("Data CSV kosong.")
            data = _parse_sheet_csv(raw, keep_cols)
            del raw
            break
        except SheetFetchError:
            raise
        except csv.Error as e:
            raise SheetFetchError(f"Format CSV tidak valid: {e}")
        except urllib.error.HTTPError as e:
            last_err = e
            if e.code in (400, 401, 403, 404):
                raise SheetFetchError(
                    f"Spreadsheet tidak bisa diakses (HTTP {e.code}). "
                    f"Cek Spreadsheet ID/GID dan izin share.")
        except (urllib.error.URLError, socket.timeout, TimeoutError,
                http.client.HTTPException, ConnectionError, OSError) as e:
            last_err = e
        if attempt < FETCH_RETRIES:
            wait = 2 * attempt
            log.warning(f"  Ambil data gagal (percobaan {attempt}/{FETCH_RETRIES}): {last_err} "
                        f"-> ulang {wait} detik lagi")
            ctl_sleep(wait)

    if data is None:
        raise SheetFetchError(f"Gagal mengambil data setelah {FETCH_RETRIES}x percobaan: {last_err}")
    return data


def write_recon_report(report_rows: list, report_path: str):
    order = [
        (STATUS_BELUM_IN,  "SKU BELUM EXCEL IN"),
        (STATUS_BELUM_OUT, "SKU BELUM EXCEL OUT"),
    ]
    by_status = {}
    for row in report_rows:
        by_status.setdefault(row["status"], []).append(row)

    lines = []
    lines.append("=" * 70)
    lines.append(f"LAPORAN REKONSILIASI SYNC RECON (Kolom BM/BN/BQ) - {STORE_CODE}")
    lines.append(f"Waktu: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 70)
    for status_key, title in order:
        items = by_status.get(status_key, [])
        lines.append("")
        lines.append(f"[{title}]  (total: {len(items)})")
        if items:
            for it in sorted(items, key=lambda r: (r["sku"], r["bn"])):
                lines.append(f"    SKU: {it['sku']:<12} No Ctn/Box: {it['bn']:<10}")
        else:
            lines.append("    (tidak ada)")
    lines.append("")
    lines.append("=" * 70)

    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    with open(report_path, 'w', encoding='utf-8', newline='') as f:
        f.write('\n'.join(lines))


def _recon_norm_sku(v: str):
    v = str(v if v is not None else "").strip().upper()
    return v or None


def reconcile_ks_master() -> dict:
    """[SYNC RECON] Rekonsiliasi berbasis status per-baris pada kolom BQ.

    Sumber data: kolom BM (SKU) ke bawah, dengan status dicek dari kolom BQ
    ke bawah (sheet & range yang sama dengan sinkronisasi biasa).

      - BQ = "SKU Belum Excel In"
            -> EXCEL IN BOX.txt   format (BN,BM)
            -> +PRINT IN LABEL.txt  format (BM,1)
      - BQ = "SKU Belum Excel Out"
            -> EXCEL OUT BOX.txt  format (,BM)
            -> -PRINT OUT LABEL.txt format (BM,1)

    Semua file output disimpan di lokasi folder yang sama dengan fungsi
    Sync IN/OUT (FILE1-FILE4).
    """
    result = {
        "ok": False, "found_in": False, "found_out": False,
        "error": None, "new_box_codes": [], "stats": {}, "report": [],
    }
    try:
        ui_log("===== SYNC RECON: REKONSILIASI BERDASARKAN STATUS (KOLOM BQ) =====")

        all_rows = fetch_sheet_rows()
        log.info(f"Total baris data terbaca: {len(all_rows)}")

        out1, out2 = [], []   # EXCEL IN BOX.txt   , +PRINT IN LABEL.txt
        out3, out4 = [], []   # EXCEL OUT BOX.txt  , -PRINT OUT LABEL.txt
        seen_in  = set()
        seen_out = set()
        report = []
        skipped = 0
        new_box_codes = set()

        for cols in all_rows:
            if len(cols) <= BQ_COL_INDEX:
                continue

            bm_raw = cols[BM_COL_INDEX].strip()
            bn_raw = cols[BN_COL_INDEX].strip()
            bq_raw = cols[BQ_COL_INDEX].strip()

            sku = _recon_norm_sku(bm_raw)
            if not sku or not bq_raw:
                if bm_raw or bq_raw:
                    skipped += 1
                continue

            if bq_raw == STATUS_BELUM_IN:
                bn_substring = bn_raw[:4]
                dedup_key = f"{bn_raw}|{sku}"
                if dedup_key in seen_in:
                    continue
                seen_in.add(dedup_key)
                out1.append(f"{bn_substring},{sku}")
                out2.append(f"{sku},1")
                if bn_raw:
                    new_box_codes.add(bn_raw)
                report.append({"status": STATUS_BELUM_IN, "sku": sku, "bn": bn_raw})

            elif bq_raw == STATUS_BELUM_OUT:
                dedup_key = f"{bn_raw}|{sku}"
                if dedup_key in seen_out:
                    continue
                seen_out.add(dedup_key)
                out3.append(f",{sku}")
                out4.append(f"{sku},3")
                report.append({"status": STATUS_BELUM_OUT, "sku": sku, "bn": bn_raw})

        if skipped:
            log.info(f"  Baris diabaikan (SKU/status kosong/tidak valid): {skipped}")

        found_in  = bool(out1)
        found_out = bool(out3)

        write_txt(FILE1, out1)
        write_txt(FILE2, out2)
        if found_in:
            ui_log(f"RECON IN  OK: {len(out1)} SKU ('{STATUS_BELUM_IN}')")
        else:
            ui_log(f"RECON IN  - tidak ada SKU '{STATUS_BELUM_IN}', FILE1 & FILE2 dikosongkan")

        write_txt(FILE3, out3)
        write_txt(FILE4, out4)
        if found_out:
            ui_log(f"RECON OUT OK: {len(out3)} SKU ('{STATUS_BELUM_OUT}')")
        else:
            ui_log(f"RECON OUT - tidak ada SKU '{STATUS_BELUM_OUT}', FILE3 & FILE4 dikosongkan")

        ui_log(
            f"RINGKASAN RECON -> Belum Excel In:{len(out1)}  "
            f"Belum Excel Out:{len(out3)}"
        )

        # Kode Brown Box baru (BN, kode penuh) dari baris "Belum Excel In", untuk cetak Brown Box
        new_box_codes = sorted(new_box_codes)

        report_path = os.path.join(DIR_PROSES, f"RECON REPORT {STORE_CODE}.txt")
        try:
            write_recon_report(report, report_path)
            ui_log(f"Laporan rekonsiliasi disimpan: {report_path}")
        except Exception as e:
            log.warning(f"Gagal menulis laporan rekonsiliasi: {e}")

        ui_log("===== SYNC RECON SELESAI =====")
        result.update(
            ok=True, found_in=found_in, found_out=found_out,
            new_box_codes=new_box_codes,
            stats={
                "BELUM_IN":  out1,
                "BELUM_OUT": out3,
            },
            report=report,
        )
        return result

    except Exception as e:
        ui_log(f"ERROR rekonsiliasi: {e}", level=logging.ERROR)
        result["error"] = str(e)
        return result


def _check_sync_guard(prev_active: int, out_count: int, total_valid: int):
    """Pengaman: cegah OUT massal palsu (mis. sheet gagal terbaca penuh, baris
    terhapus/tergeser). Bisa dilewati 1x dengan membuat file force_sync.flag."""
    reason = None
    if prev_active and total_valid == 0:
        reason = "spreadsheet terbaca KOSONG padahal sebelumnya ada data"
    elif prev_active >= GUARD_MIN_ACTIVE and out_count > prev_active * GUARD_OUT_RATIO:
        reason = (f"{out_count} dari {prev_active} data aktif akan dianggap OUT sekaligus "
                  f"(> {int(GUARD_OUT_RATIO * 100)}%)")
    if not reason:
        return
    if FORCE_SYNC_FLAG.exists():
        log.warning(f"Pengaman dilewati lewat {FORCE_SYNC_FLAG.name}: {reason}")
        try:
            FORCE_SYNC_FLAG.unlink()
        except Exception:
            pass
        return
    raise Exception(
        f"SYNC DIBATALKAN (pengaman): {reason}.\n"
        "Kemungkinan spreadsheet gagal terbaca penuh / banyak baris terhapus. "
        "Tidak ada data yang diubah.\n"
        f"Jika memang benar, buat file kosong bernama '{FORCE_SYNC_FLAG.name}' di folder:\n"
        f"{APP_DIR}\nlalu klik Sync lagi.")


def _build_plan(all_rows: list, state) -> dict:
    """Bandingkan isi sheet dengan state -> rencana IN/OUT (TANPA menulis apa pun)."""
    stats = {"IN": [], "RE_IN": [], "MOVE": [], "ACTIVE": [],
             "OUT": [], "DUPLICATE": [], "INVALID": []}
    active_keys, dead_keys = state.active, state.dead

    seen_count = {}
    valid_rows = []
    for cols in all_rows:
        if len(cols) <= C_COL_INDEX:
            continue

        b_val = cols[B_COL_INDEX]
        c_val = cols[C_COL_INDEX]
        ah_val = cols[AH_COL_INDEX] if len(cols) > AH_COL_INDEX else ""

        if not b_val and not c_val:
            continue

        if not b_val or not c_val or not c_val.isdigit():
            stats["INVALID"].append(f"{b_val or '(kosong)'}|{c_val or '(kosong)'}")
            continue

        key = make_key(b_val, c_val)
        seen_count[key] = seen_count.get(key, 0) + 1
        valid_rows.append((b_val, c_val, ah_val, key))

    dup_keys = {k for k, cnt in seen_count.items() if cnt > 1}
    stats["DUPLICATE"] = sorted(dup_keys)

    # Semua kunci yang MASIH ADA di sheet (termasuk duplikat). Duplikat tidak boleh
    # dianggap "hilang" -> kalau tidak, data aktif yang kebetulan terduplikasi bisa
    # salah dihitung OUT.
    present_keys = set(seen_count)

    new_active_keys = set()
    row_by_key = {}
    sku_new_map = {}
    family_map = {}
    ah_values_in_sheet = set()

    for b_val, c_val, ah_val, key in valid_rows:
        sku_new_map.setdefault(c_val, set()).add(b_val)
        fam = (b_base(b_val), c_val) if FAMILY_MATCH_SKU else b_base(b_val)
        family_map.setdefault(fam, set()).add(b_val)
        if ah_val:
            ah_values_in_sheet.add(ah_val)
        if key in dup_keys:
            continue          # duplikat baru: tidak di-IN sampai dibersihkan
        new_active_keys.add(key)
        row_by_key[key] = (b_val, c_val, ah_val)

    # ---- IN ----
    in_items = []
    for key in sorted(new_active_keys):
        if key in active_keys:
            continue
        b_val, c_val, _ah = row_by_key[key]
        stats["RE_IN" if key in dead_keys else "IN"].append(key)
        in_items.append({
            "box": f"{b_val[:4]},{c_val}", "label": f"{c_val},1",
            "keys": [key], "box_code": b_val,
        })

    # ---- OUT / MOVE / ACTIVE ----
    out_items, out_by_sku = [], {}
    for key in sorted(active_keys - present_keys):
        if "|" not in key:
            continue
        b_from_key, c_from_key = key.rsplit("|", 1)
        base = b_base(b_from_key)

        if c_from_key in ah_values_in_sheet:
            stats["ACTIVE"].append(key)
            continue
        if sku_new_map.get(c_from_key):
            stats["MOVE"].append(key)
            continue
        if family_map.get((base, c_from_key) if FAMILY_MATCH_SKU else base):
            stats["ACTIVE"].append(key)
            continue

        item = out_by_sku.get(c_from_key)
        if item is None:
            item = {"box": f",{c_from_key}", "label": f"{c_from_key},3", "keys": []}
            out_by_sku[c_from_key] = item
            out_items.append(item)
        item["keys"].append(key)
        stats["OUT"].append(key)

    return {"stats": stats, "in_items": in_items, "out_items": out_items,
            "present_keys": present_keys, "valid_count": len(valid_rows)}


def _merge_stable(p1: dict, p2: dict):
    """Gabungkan dua pembacaan sheet berurutan: hanya perubahan yang SAMA di kedua
    pembacaan yang diproses. Perubahan yang muncul/hilang di antara dua pembacaan
    (mis. staf sedang cut-paste baris) ditunda ke Sync berikutnya, bukan dianggap
    IN/OUT. Return (plan_stabil, jumlah_ditunda)."""
    in1 = {it["keys"][0] for it in p1["in_items"]}
    in2 = {it["keys"][0] for it in p2["in_items"]}
    in_items = [it for it in p2["in_items"] if it["keys"][0] in in1]
    unstable = len(in1 ^ in2)

    g1 = {it["box"]: frozenset(it["keys"]) for it in p1["out_items"]}
    g2 = {it["box"]: frozenset(it["keys"]) for it in p2["out_items"]}
    out_items = [it for it in p2["out_items"] if g1.get(it["box"]) == frozenset(it["keys"])]
    unstable += sum(1 for b in set(g1) | set(g2) if g1.get(b) != g2.get(b))

    stable_in = {it["keys"][0] for it in in_items}
    stable_out = {k for it in out_items for k in it["keys"]}
    stats = dict(p2["stats"])
    stats["IN"] = [k for k in stats["IN"] if k in stable_in]
    stats["RE_IN"] = [k for k in stats["RE_IN"] if k in stable_in]
    stats["OUT"] = [k for k in stats["OUT"] if k in stable_out]

    plan = dict(p2)
    plan.update(stats=stats, in_items=in_items, out_items=out_items)
    return plan, unstable


def sync_from_sheet() -> dict:
    """Bandingkan sheet dengan state, hasilkan item IN/OUT.
    TIDAK mengubah state di sini: state baru di-commit per batch setelah otomasi
    EBarcode sukses, lihat process_kind().
    Alur: baca sheet -> rencana -> (jika ada perubahan) baca ulang & ambil yang
    stabil saja -> pengaman OUT massal -> tulis file."""
    result = {
        "ok": False, "found_new": False, "found_out": False,
        "error": None, "new_box_codes": [], "stats": {},
        "in_items": [], "out_items": [], "state": None, "baseline": None, "unstable": 0,
    }
    state, rows_read, prev_active = None, 0, 0
    try:
        ui_log("===== SINKRONISASI DATA (Google Sheets) =====")

        rows = fetch_sheet_rows()
        rows_read = len(rows)
        log.info(f"Total baris data terbaca: {rows_read}")

        state = SyncState.load()
        prev_active = len(state.active)
        log.info(f"Snapshot lama: {prev_active} kunci aktif, "
                 f"{len(state.dead)} kunci pernah OUT (riwayat)")

        plan = _build_plan(rows, state)
        del rows

        # ---- PEMAKAIAN PERTAMA: catat semua data yang ada sebagai baseline ----
        if state.is_new and FIRST_RUN_BASELINE and plan["present_keys"]:
            state.baseline(plan["present_keys"])
            n = len(plan["present_keys"])
            ui_log(f"PERTAMA KALI: {n} data yang sudah ada dicatat sebagai baseline "
                   f"(TIDAK dicetak). Sync berikutnya hanya memproses data baru/keluar.")
            state.log_sync("BASELINE", rows=rows_read, active_before=0, n_stay=n)
            result.update(ok=True, baseline=n, stats=plan["stats"], state=state)
            return result

        # ---- VERIFIKASI: baca ulang bila ada perubahan, ambil yang stabil saja ----
        unstable = 0
        if (plan["in_items"] or plan["out_items"]) and STABLE_READ_DELAY > 0:
            ui_log(f"Perubahan terdeteksi, verifikasi dengan pembacaan kedua "
                   f"({STABLE_READ_DELAY:g} detik)...")
            ctl_sleep(STABLE_READ_DELAY)
            rows2 = fetch_sheet_rows()
            plan2 = _build_plan(rows2, state)
            del rows2
            plan, unstable = _merge_stable(plan, plan2)
            if unstable:
                ui_log(f"PERHATIAN: {unstable} perubahan belum stabil (berubah di antara dua "
                       f"pembacaan, mungkin sedang diedit). Ditunda ke Sync berikutnya.",
                       level=logging.WARNING)

        stats = plan["stats"]
        in_items, out_items = plan["in_items"], plan["out_items"]

        for k in stats["DUPLICATE"]:
            log.info(f"  [DUPLICATE] {k}")
        for k in stats["OUT"]:
            log.info(f"  [OUT] {k} -> tidak ada relasi keluarga aktif tersisa")

        # ---- Pengaman sebelum menulis apa pun ----
        _check_sync_guard(prev_active, len(stats["OUT"]), plan["valid_count"])

        found_new = bool(in_items)
        found_out = bool(out_items)

        if found_new:
            write_txt(FILE1, [i["box"] for i in in_items])
            write_txt(FILE2, [i["label"] for i in in_items])
            ui_log(f"IN  OK: {len(in_items)} baris "
                   f"(baru: {len(stats['IN'])}, RE-IN: {len(stats['RE_IN'])})")
        else:
            ui_log("IN  - tidak ada data baru, file dibiarkan")

        write_txt(FILE3, [i["box"] for i in out_items])
        write_txt(FILE4, [i["label"] for i in out_items])
        if found_out:
            ui_log(f"OUT OK: {len(out_items)} baris (kunci OUT: {len(stats['OUT'])})")
        else:
            ui_log("OUT - tidak ada SKU yang benar-benar keluar, FILE3 & FILE4 dikosongkan")

        ui_log(
            f"RINGKASAN -> IN:{len(stats['IN'])}  RE-IN:{len(stats['RE_IN'])}  "
            f"MOVE:{len(stats['MOVE'])}  ACTIVE(tetap):{len(stats['ACTIVE'])}  "
            f"OUT:{len(stats['OUT'])}  DUPLICATE:{len(stats['DUPLICATE'])}  "
            f"INVALID:{len(stats['INVALID'])}  DITUNDA:{unstable}"
        )
        if stats["DUPLICATE"]:
            ui_log(f"PERHATIAN: {len(stats['DUPLICATE'])} data DUPLICATE ditemukan, "
                   f"tidak di-IN sampai dibersihkan di spreadsheet.", level=logging.WARNING)
        if stats["INVALID"]:
            ui_log(f"PERHATIAN: {len(stats['INVALID'])} baris INVALID (No Ctn/SKU kosong "
                   f"atau format salah) diabaikan.", level=logging.WARNING)

        state.log_sync(
            "OK", rows=rows_read, active_before=prev_active,
            n_in=len(stats["IN"]) + len(stats["RE_IN"]), n_out=len(stats["OUT"]),
            n_move=len(stats["MOVE"]), n_stay=len(stats["ACTIVE"]),
            n_dup=len(stats["DUPLICATE"]), n_invalid=len(stats["INVALID"]), n_unstable=unstable)

        report_path = os.path.join(DIR_PROSES, f"STATUS REPORT {STORE_CODE}.txt")
        try:
            write_status_report(stats, report_path, history=state.history(10))
            ui_log(f"Laporan status lengkap disimpan: {report_path}")
        except Exception as e:
            log.warning(f"Gagal menulis laporan status: {e}")

        ui_log("===== SINKRONISASI SELESAI (state disimpan setelah cetak sukses) =====")
        result.update(
            ok=True, found_new=found_new, found_out=found_out, unstable=unstable,
            new_box_codes=sorted({i["box_code"] for i in in_items}),
            stats=stats, in_items=in_items, out_items=out_items, state=state,
        )
        return result

    except Exception as e:
        ui_log(f"ERROR sinkronisasi: {e}", level=logging.ERROR)
        result["error"] = str(e)
        if state is not None:
            state.log_sync("ERROR", note=str(e).replace("\n", " "),
                           rows=rows_read, active_before=prev_active)
        return result


# ============================================================
# BAGIAN 2: OTOMASI EBARCODE (dari PrintAuto.py)
# ============================================================

def create_automation_desktop() -> bool:
    global _automation_desktop, _user_desktop
    if _automation_desktop is not None:
        destroy_automation_desktop()   # jangan menumpuk desktop sisa proses sebelumnya
    if not PYVDA_OK:
        log.warning("pyvda tidak terinstall — jalankan: pip install pyvda")
        log.warning("Otomasi berjalan di desktop aktif (bisa mengganggu user).")
        return False
    try:
        _user_desktop = VirtualDesktop.current()
        _automation_desktop = VirtualDesktop.create()
        log.info(f"Virtual Desktop otomasi dibuat di background (#{_automation_desktop.number}), "
                 f"user tetap di #{_user_desktop.number}")
        return True
    except Exception as e:
        log.error(f"Gagal buat Virtual Desktop: {e}")
        return False


def destroy_automation_desktop():
    global _automation_desktop, _user_desktop
    if not PYVDA_OK or _automation_desktop is None:
        return
    try:
        _automation_desktop.remove()
        log.info("Desktop otomasi dihapus.")
    except Exception as e:
        log.error(f"Gagal hapus desktop otomasi: {e}")

    _automation_desktop = None
    _user_desktop = None


def switch_to_user_desktop():
    pass


def wait_hwnd_ready(hwnd: int, timeout: float = 5.0) -> bool:
    user32 = ctypes.windll.user32
    start = ctl_time()
    while ctl_time() - start < timeout:
        if (user32.IsWindow(hwnd) and
                user32.IsWindowVisible(hwnd) and
                user32.SendMessageTimeoutW(hwnd, 0, 0, 0, 0x0002, 500, None) != 0):
            return True
        ctl_sleep(0.1)
    return False


def move_window_to_automation_desktop(hwnd: int):
    if not PYVDA_OK or _automation_desktop is None:
        return
    if not wait_hwnd_ready(hwnd):
        log.warning(f"  HWND {hwnd} tidak siap dalam timeout, skip pindah ke vdesk.")
        return
    try:
        AppView(hwnd).move(_automation_desktop)
        log.info(f"  Window {hwnd} dipindah ke desktop otomasi.")
    except Exception as e:
        log.warning(f"  Gagal pindah window {hwnd}: {e}")


def wait_win(class_name, timeout=TIMEOUT):
    start = ctl_time()
    while ctl_time() - start < timeout:
        wins = Desktop(backend="win32").windows(class_name=class_name)
        if wins:
            return wins[0]
        ctl_sleep(0.25)
    return None


def get_app():
    return Application(backend="win32").connect(path=EXE_PATH)


def kill_ebarcode():
    """Paksa tutup semua window EBarcode via WM_SYSCOMMAND/SC_CLOSE
    (bukan Alt+F4/send_keys), lalu fallback ke taskkill jika masih
    ada proses berjalan. Tetap bekerja walau PC locked."""
    log.warning("  [KILL] Menutup semua window EBarcode...")

    for cls in ["TdlgData", "TdlgStoreSHELF", "TMessageForm", "TfrmPrint", "TfrmBarcode"]:
        for attempt in range(3):
            wins = Desktop(backend="win32").windows(class_name=cls)
            if not wins:
                break
            for w in wins:
                try:
                    safe_set_focus(w)
                    ctl_sleep(0.05)
                    close_window_message(w.handle)
                    ctl_sleep(0.15)
                    log.warning(f"  [KILL] Close message ke {cls}")
                except Exception as e:
                    log.warning(f"  [KILL] Gagal close {cls}: {e}")
            ctl_sleep(0.25)

    ctl_sleep(0.75)
    wins = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if wins:
        log.warning("  [KILL] EBarcode masih ada, paksa kill via taskkill...")
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", "EBarcode.exe"],
                capture_output=True, timeout=5
            )
            ctl_sleep(0.75)
            log.warning("  [KILL] taskkill selesai")
        except Exception as e:
            log.warning(f"  [KILL] taskkill error: {e}")
    else:
        log.warning("  [KILL] Semua window EBarcode sudah tertutup")


def fail(msg: str) -> bool:
    ui_log(f"*** FAIL: {msg}", level=logging.ERROR)
    kill_ebarcode()
    switch_to_user_desktop()
    return False


def run_print_flow(trigger_file: Path, label_search: str) -> bool:
    """Pembungkus: apa pun hasilnya (sukses/gagal/exception), Virtual Desktop
    otomasi SELALU dibersihkan. Dulu jalur gagal tidak membersihkan sehingga
    desktop menumpuk seiring lama pemakaian."""
    try:
        return _run_print_flow_impl(trigger_file, label_search)
    except StopRequested:
        ui_log("*** DIHENTIKAN manual oleh pengguna.", level=logging.WARNING)
        if STOP_CLOSES_EBARCODE:
            with control.cleanup():
                try:
                    kill_ebarcode()
                except Exception:
                    pass
        return False
    except Exception as e:
        ui_log(f"*** ERROR tak terduga di otomasi EBarcode: {e}", level=logging.ERROR)
        log.exception("Detail error otomasi EBarcode:")
        try:
            kill_ebarcode()
        except Exception:
            pass
        return False
    finally:
        destroy_automation_desktop()


def _run_print_flow_impl(trigger_file: Path, label_search: str) -> bool:
    """
    Jalankan otomasi EBarcode.
    - trigger_file : file yang isinya di-paste ke TRichEdit1 (data Excel In).
    - label_search : teks yang dicari (case-insensitive, substring) di
                     TCheckListBox2 pada STEP 7 (mis. "PRINT IN LABEL" atau
                     "PRINT OUT LABEL") untuk dicentang sebelum print.
    """
    ui_log(f"=== START EBarcode | {trigger_file.name} | label='{label_search}' ===")

    try:
        n_lines = len([l for l in trigger_file.read_text(encoding="utf-8", errors="ignore").splitlines()
                       if l.strip()])
    except Exception:
        n_lines = 0

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if leftover:
        log.warning("  EBarcode masih terbuka dari sesi sebelumnya, tutup dulu...")
        kill_ebarcode()
        ctl_sleep(0.75)

    using_vdesk = create_automation_desktop()
    if using_vdesk:
        log.info("Otomasi berjalan di Virtual Desktop tersembunyi — user tidak terganggu.")

    # STEP 1: Buka EBarcode
    ui_log("STEP 1 - Buka EBarcode")
    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        ctl_sleep(1.5)

    if not wait_win("TfrmBarcode"):
        return fail("TfrmBarcode tidak muncul")

    app = get_app()
    frm = app.window(class_name="TfrmBarcode")

    if using_vdesk:
        move_window_to_automation_desktop(frm.handle)
        ctl_sleep(0.15)

    safe_set_focus(frm)
    ctl_sleep(0.35)
    log.info("  OK: TfrmBarcode aktif")

    # STEP 2: Klik Print (TButton3)
    ui_log("STEP 2 - Klik Print (TButton3)")
    try:
        btn_print = frm.child_window(class_name="TButton", found_index=2)
        if not click_control_message(btn_print):
            return fail("Gagal klik TButton3 (safe click gagal)")
        ctl_sleep(0.5)
        log.info("  OK")
    except Exception as e:
        return fail(f"Gagal klik TButton3: {e}")

    # STEP 3: TfrmPrint -> klik tab BOX (tepat di sebelah kanan tab SHELF)
    ui_log("STEP 3 - Klik tab BOX di TfrmPrint")
    if not wait_win("TfrmPrint"):
        return fail("TfrmPrint tidak muncul")

    app = get_app()
    frm_print = app.window(class_name="TfrmPrint")

    if using_vdesk:
        move_window_to_automation_desktop(frm_print.handle)
        ctl_sleep(0.15)

    safe_set_focus(frm_print)
    ctl_sleep(0.25)

    try:
        tab_ctrl  = frm_print.child_window(class_name="TPageControl", found_index=0)
        tab_texts = []
        for child in frm_print.children():
            try:
                if child.class_name() == "TTabSheet":
                    tab_texts.append(child.window_text().strip())
            except Exception:
                pass
        log.info(f"  Tab list: {tab_texts}")

        # SetForegroundWindow dilewati (gagal saat PC locked) — tidak
        # diperlukan karena klik di bawah ini sudah berbasis PostMessage.
        ctl_sleep(0.15)

        TCM_GETCURSEL   = 0x130B
        TCM_SETCURFOCUS = 0x1330  # set tab by index + otomatis kirim notifikasi TCN_SELCHANGE

        sel_before = ctypes.windll.user32.SendMessageW(tab_ctrl.handle, TCM_GETCURSEL, 0, 0)

        box_tab = None
        for child in tab_ctrl.children():
            try:
                text_upper = child.window_text().strip().upper()
                if text_upper == "BOX" or ("BOX" in text_upper and "SHELF" not in text_upper):
                    box_tab = child
                    break
            except Exception:
                pass
        if box_tab:
            click_control_message(box_tab)
        else:
            # Fallback koordinat relatif jika teks "BOX" tidak ditemukan
            # (BOX berada tepat di kanan SHELF pada TPageControl yang sama)
            click_message(tab_ctrl.handle, 110, 8)
        ctl_sleep(0.25)

        sel_after = ctypes.windll.user32.SendMessageW(tab_ctrl.handle, TCM_GETCURSEL, 0, 0)
        log.info(f"  [DIAG] Tab selection sebelum={sel_before} sesudah={sel_after} (tab_texts={tab_texts})")

        if sel_before == sel_after:
            # Klik tidak mengubah tab. Coba paksa pakai TCM_SETCURFOCUS,
            # menebak index tab BOX dari daftar teks tab yang terlihat.
            guess_index = None
            for idx, t in enumerate(tab_texts):
                tu = t.strip().upper()
                if tu == "BOX" or ("BOX" in tu and "SHELF" not in tu):
                    guess_index = idx
                    break
            if guess_index is None:
                guess_index = len(tab_texts) - 1 if tab_texts else 0
            ctypes.windll.user32.SendMessageW(tab_ctrl.handle, TCM_SETCURFOCUS, guess_index, 0)
            ctl_sleep(0.15)
            sel_forced = ctypes.windll.user32.SendMessageW(tab_ctrl.handle, TCM_GETCURSEL, 0, 0)
            log.info(f"  [DIAG] Klik tidak mengubah tab, dicoba TCM_SETCURFOCUS index={guess_index} -> selection={sel_forced}")

        box_tab_sheet = None
        for child in frm_print.descendants():
            try:
                text_upper = child.window_text().strip().upper()
                if child.class_name() == "TTabSheet" and (
                    text_upper == "BOX" or ("BOX" in text_upper and "SHELF" not in text_upper)
                ):
                    box_tab_sheet = child
                    break
            except Exception:
                pass

        nav = None
        search_root = box_tab_sheet if box_tab_sheet else frm_print
        for ctrl in search_root.descendants():
            try:
                if ctrl.class_name() == "TDBNavigator":
                    nav = ctrl
                    break
            except Exception:
                pass

        if nav is None:
            return fail("TDBNavigator tidak ditemukan di tab BOX")

        grid = None
        for ctrl in search_root.descendants():
            try:
                if ctrl.class_name() == "TwwDBGrid":
                    grid = ctrl
                    break
            except Exception:
                pass

        nav_rect  = nav.rectangle()
        grid_rect = grid.rectangle() if grid else None
        log.info(f"  nav_rect={nav_rect} grid_rect={grid_rect}")

        clicked_ok = click_excel_in_with_retry(frm_print, nav_rect, grid_rect)
        if not clicked_ok:
            return fail("Klik Excel In tidak terdeteksi berhasil di titik manapun")

        ctl_sleep(0.25)
        log.info("  OK: Tab BOX diklik dan Excel In diklik")
    except Exception as e:
        return fail(f"Gagal klik tab BOX: {e}")

    # STEP 3c
    ui_log("STEP 3c - Tunggu popup konfirmasi setelah Excel In di navigator")
    popup_3c = wait_win("TMessageForm", timeout=0.5)
    if popup_3c:
        try:
            app = get_app()
            dlg = app.window(class_name="TMessageForm")
            if using_vdesk:
                move_window_to_automation_desktop(dlg.handle)
            log.info(f"  Popup muncul: {dlg.window_text().strip()!r}")
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            log.info("  OK: Popup konfirmasi diklik")
            ctl_sleep(0.25)
            for _ in range(20):
                wins = Desktop(backend="win32").windows(class_name="TMessageForm")
                if not wins:
                    break
                ctl_sleep(0.15)
        except Exception as e:
            log.info(f"  Popup konfirmasi error: {e}, lanjut")
    else:
        log.info("  Tidak ada popup konfirmasi, lanjut")

    # STEP 4: TdlgStoreSHELF -> paste ke TRichEdit1
    ui_log("STEP 4 - Paste path file ke TRichEdit1")
    if not wait_win("TdlgStoreSHELF"):
        return fail("TdlgStoreSHELF tidak muncul")

    app = get_app()
    shelf_dlg = app.window(class_name="TdlgStoreSHELF")

    if using_vdesk:
        move_window_to_automation_desktop(shelf_dlg.handle)
        ctl_sleep(0.15)

    safe_set_focus(shelf_dlg)
    ctl_sleep(0.25)

    try:
        rich_edit = shelf_dlg.child_window(class_name="TRichEdit", found_index=0)
        safe_set_focus(rich_edit)
        ctl_sleep(0.1)
        file_content = trigger_file.read_text(encoding="utf-8", errors="ignore").strip()
        log.info(f"  Isi file ({len(file_content)} chars):\n{file_content[:200]}")
        rich_edit.set_edit_text(file_content)
        ctl_sleep(0.15)
        log.info("  OK: Konten file di-paste ke TRichEdit1")
    except Exception as e:
        return fail(f"Gagal paste ke TRichEdit1: {e}")

    # STEP 5: Klik Excel In
    ui_log("STEP 5 - Klik Excel In (TcxButton2)")
    try:
        btn_excel_in = shelf_dlg.child_window(class_name="TcxButton", found_index=1)
        if not click_control_message(btn_excel_in):
            return fail("Gagal klik TcxButton2 (Excel In) (safe click gagal)")
        ctl_sleep(0.5)
        log.info("  OK: Excel In diklik")
    except Exception as e:
        return fail(f"Gagal klik TcxButton2 (Excel In): {e}")

    ui_log("STEP 5b - Tunggu popup konfirmasi setelah TcxButton2 (Excel In), lalu klik")
    msg_win = wait_win("TMessageForm", timeout=max(30, 10 + 0.5 * n_lines))  # skala dgn jumlah baris
    if msg_win:
        try:
            app = get_app()
            dlg = app.window(class_name="TMessageForm")
            if using_vdesk:
                move_window_to_automation_desktop(dlg.handle)
            log.info(f"  Popup muncul: {dlg.window_text().strip()!r}")
            if not click_control_message(dlg.child_window(class_name="TButton", found_index=0)):
                return fail("Gagal klik OK TMessageForm (safe click gagal)")
            log.info("  OK: Popup diklik, tunggu hilang...")
            for _ in range(20):
                wins = Desktop(backend="win32").windows(class_name="TMessageForm")
                if not wins:
                    log.info("  OK: Popup sudah tertutup")
                    break
                ctl_sleep(0.15)
        except Exception as e:
            return fail(f"Gagal klik OK TMessageForm: {e}")
    else:
        log.info("  TMessageForm tidak muncul, lanjut")

    ui_log("STEP 5c - Klik Close TdlgStoreSHELF")
    try:
        app = get_app()
        shelf_win = app.window(class_name="TdlgStoreSHELF")
        safe_set_focus(shelf_win)
        ctl_sleep(0.1)
        close_window_message(shelf_win.handle)
        log.info("  OK: Close message dikirim ke TdlgStoreSHELF, tunggu hilang...")
        for _ in range(20):
            wins = Desktop(backend="win32").windows(class_name="TdlgStoreSHELF")
            if not wins:
                log.info("  OK: TdlgStoreSHELF sudah tertutup")
                break
            ctl_sleep(0.15)
    except Exception as e:
        log.info(f"  Close TdlgStoreSHELF error: {e}, lanjut")

    # STEP 6: TfrmPrint -> klik Zebex
    ui_log("STEP 6 - TfrmPrint -> klik Zebex (TcxButton9)")
    if not wait_win("TfrmPrint", timeout=10):
        return fail("TfrmPrint tidak muncul setelah Excel In")

    app = get_app()
    frm_print2 = app.window(class_name="TfrmPrint")
    if using_vdesk:
        move_window_to_automation_desktop(frm_print2.handle)
        ctl_sleep(0.15)
    safe_set_focus(frm_print2)
    ctl_sleep(0.25)

    try:
        zebex = None
        for i, btn in enumerate(frm_print2.children(class_name="TcxButton")):
            log.info(f"  TcxButton[{i}]: '{btn.window_text().strip()}'")
            if i == 8:
                zebex = btn
                break
        if not zebex:
            try:
                zebex = frm_print2.child_window(title="Zebex", class_name="TcxButton")
            except Exception:
                zebex = frm_print2.child_window(class_name="TcxButton", found_index=8)
        if not click_control_message(zebex):
            return fail("Gagal klik TcxButton9 (Zebex) (safe click gagal)")
        ctl_sleep(0.4)
        log.info("  OK: Zebex diklik")
    except Exception as e:
        return fail(f"Gagal klik TcxButton9 (Zebex): {e}")

    # STEP 6b: Tunggu & tutup 'Zebex Download Utility' jika muncul
    ui_log("STEP 6b - Tunggu & tutup 'Zebex Download Utility' jika muncul")

    ZEBEX_DL_TIMEOUT = 3

    def find_zebex_download_window():
        try:
            wins = Desktop(backend="win32").windows(class_name="TfrmMain")
        except Exception:
            wins = []
        for w in wins:
            try:
                if "ZEBEX DOWNLOAD" in w.window_text().upper():
                    return w
            except Exception:
                pass
        return None

    zebex_dl_win = None
    start = ctl_time()
    while ctl_time() - start < ZEBEX_DL_TIMEOUT:
        zebex_dl_win = find_zebex_download_window()
        if zebex_dl_win:
            break
        ctl_sleep(0.25)

    if zebex_dl_win:
        try:
            log.info(f"  'Zebex Download Utility' terdeteksi: "
                     f"{zebex_dl_win.window_text().strip()!r} -> ditutup")
            safe_set_focus(zebex_dl_win)
            ctl_sleep(0.15)

            close_btn = None
            for ctrl in zebex_dl_win.descendants():
                try:
                    if ctrl.window_text().strip().upper() == "CLOSE":
                        close_btn = ctrl
                        break
                except Exception:
                    pass

            if close_btn:
                click_control_message(close_btn)
                log.info("  OK: Tombol 'Close' diklik")
            else:
                close_window_message(zebex_dl_win.handle)
                log.info("  OK: Close message dikirim (tombol 'Close' tidak ditemukan)")

            ctl_sleep(0.25)
            for _ in range(20):
                if find_zebex_download_window() is None:
                    ui_log("  OK: 'Zebex Download Utility' sudah tertutup")
                    break
                ctl_sleep(0.15)
            else:
                try:
                    close_window_message(zebex_dl_win.handle)
                    ctl_sleep(0.25)
                except Exception:
                    pass
        except Exception as e:
            log.warning(f"  Gagal tutup 'Zebex Download Utility': {e}, lanjut saja")
    else:
        log.info("  'Zebex Download Utility' tidak muncul, lanjut")

    # STEP 7: TdlgData -> Un-tick -> centang label sesuai alur (IN/OUT)
    ui_log(f"STEP 7 - Dialog TdlgData (cari label: '{label_search}')")

    tdlg_found = False
    for _ in range(30):
        wins = Desktop(backend="win32").windows()
        for w in wins:
            try:
                cls = w.class_name()
                txt = w.window_text().strip()
                if cls == "TdlgData":
                    tdlg_found = True
                    break
                if cls in ("TMessageForm", "#32770", "TForm") and txt:
                    log.info(f"  [POPUP] class={cls!r} text={txt!r} — klik OK/Close")
                    try:
                        pop = Desktop(backend="win32").window(class_name=cls, title=txt)
                        click_control_message(pop.child_window(class_name="TButton", found_index=0))
                    except Exception:
                        pass
                    ctl_sleep(0.15)
            except Exception:
                pass
        if tdlg_found:
            break
        ctl_sleep(0.25)

    if not tdlg_found:
        return fail("TdlgData tidak muncul")

    app = get_app()
    data_win = app.window(class_name="TdlgData")
    if using_vdesk:
        move_window_to_automation_desktop(data_win.handle)
        ctl_sleep(0.15)
    safe_set_focus(data_win)
    ctl_sleep(0.3)

    try:
        btn_untick = data_win.child_window(class_name="TBitBtn", found_index=3)
        if not click_control_message(btn_untick):
            return fail("Gagal klik Un-tick semua (TBitBtn) (safe click gagal)")
        ctl_sleep(0.2)
        log.info("  Un-tick semua selesai")

        checklist_focus = data_win.child_window(class_name="TCheckListBox", found_index=1)
        safe_set_focus(checklist_focus)
        ctl_sleep(0.15)

        LB_GETCOUNT  = 0x018B
        LB_GETTEXT   = 0x0189
        LB_SETCURSEL = 0x0186

        checklist = data_win.child_window(class_name="TCheckListBox", found_index=1)
        hwnd  = checklist.handle
        count = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCOUNT, 0, 0)
        log.info(f"  Jumlah item TCheckListBox2: {count}")

        LB_SETCHECK  = 0x0191
        LB_GETCHECK  = 0x0190

        found = False
        for i in range(count):
            buf = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.SendMessageW(hwnd, LB_GETTEXT, i, buf)
            item_text = buf.value.strip()
            log.info(f"    [{i}] '{item_text}'")
            if label_search.upper() in item_text.upper():
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCURSEL, i, 0)
                ctl_sleep(0.075)
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCHECK, i, 1)
                ctl_sleep(0.1)
                checked = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCHECK, i, 0)
                if checked == 1:
                    log.info(f"  OK: '{item_text}' berhasil dicentang (index {i})")
                else:
                    log.warning(f"  LB_SETCHECK tidak bekerja (checked={checked}), coba klik area checkbox")
                    LB_GETITEMHEIGHT = 0x01A1
                    item_height = ctypes.windll.user32.SendMessageW(hwnd, LB_GETITEMHEIGHT, 0, 0)
                    if item_height <= 0:
                        item_height = 16
                    click_y = (i * item_height) + (item_height // 2)
                    log.info(f"  Klik checkbox: item_height={item_height} click_y={click_y}")
                    click_message(hwnd, 8, click_y)
                    ctl_sleep(0.1)
                found = True
                break

        if not found:
            return fail(f"'{label_search}' tidak ditemukan di TCheckListBox2")
        ctl_sleep(0.15)

    except Exception as e:
        return fail(f"Error step 7: {e}")

    # STEP 8: Extract >>
    ui_log("STEP 8 - Extract >> (TBitBtn7)")
    try:
        app = get_app()
        btn_extract = app.window(class_name="TdlgData").child_window(
            class_name="TBitBtn", found_index=6)
        if not click_control_message(btn_extract):
            return fail("Gagal klik TBitBtn7 (Extract) (safe click gagal)")
        ctl_sleep(1.25)
        log.info("  OK: Extract selesai")
    except Exception as e:
        return fail(f"Gagal klik TBitBtn7 (Extract): {e}")

    # STEP 9: Print >>
    ui_log("STEP 9 - Print >> (TBitBtn6)")
    try:
        app = get_app()
        btn_print2 = app.window(class_name="TdlgData").child_window(
            class_name="TBitBtn", found_index=5)
        if not click_control_message(btn_print2):
            return fail("Gagal klik TBitBtn6 (Print) (safe click gagal)")
        ctl_sleep(1.25)
        log.info("  OK: Print dikirim")
    except Exception as e:
        return fail(f"Gagal klik TBitBtn6 (Print): {e}")

    ui_log("STEP 9b - Klik TButton1 (close popup setelah print)")
    popup = wait_win("TMessageForm", timeout=500)
    if popup:
        try:
            app = get_app()
            dlg = app.window(class_name="TMessageForm")
            if using_vdesk:
                move_window_to_automation_desktop(dlg.handle)
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            ctl_sleep(0.25)
            log.info("  OK: Popup print ditutup")
        except Exception as e:
            log.info(f"  TMessageForm TButton1 error: {e}")
    else:
        log.info("  Tidak ada popup, lanjut")

    ui_log("STEP 9c - Close TdlgData (TBitBtn8)")
    try:
        app = get_app()
        btn_close_data = app.window(class_name="TdlgData").child_window(
            class_name="TBitBtn", found_index=7)
        click_control_message(btn_close_data)
        ctl_sleep(0.25)
        log.info("  OK: TdlgData ditutup")
    except Exception as e:
        log.info(f"  TBitBtn8 error: {e}")

    if using_vdesk:
        destroy_automation_desktop()

    ui_log(f"=== SELESAI | {trigger_file.name} ===")
    return True


def open_brown_box_print_form(using_vdesk: bool):
    """
    Buka EBarcode (kalau belum jalan) -> di TfrmBarcode pilih 'BROWN BOX' di
    Select Label -> klik Print (TButton3) -> tunggu TfrmPrint muncul.
    Return: frm_print (pywinauto window) kalau berhasil, None kalau gagal.
    """
    ui_log("  BROWN BOX - Buka EBarcode")

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    leftover_print = Desktop(backend="win32").windows(class_name="TfrmPrint")
    if leftover or leftover_print:
        log.warning("  BROWN BOX - EBarcode/TfrmPrint masih terbuka dari sesi "
                    "sebelumnya, tutup dulu supaya tidak salah pilih menu...")
        kill_ebarcode()
        ctl_sleep(0.75)

    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        ctl_sleep(1.5)

    if not wait_win("TfrmBarcode", timeout=TIMEOUT):
        log.warning("  TfrmBarcode tidak muncul, lewati print BROWN BOX")
        return None

    app = get_app()
    frm = app.window(class_name="TfrmBarcode")
    if using_vdesk:
        move_window_to_automation_desktop(frm.handle)
        ctl_sleep(0.15)
    safe_set_focus(frm)
    ctl_sleep(0.25)

    log.info(f"  BROWN BOX - Pilih label '{BROWN_BOX_LABEL_NAME}' di Select Label")
    try:
        # Panah bawah saat combo TERTUTUP cuma mengganti teks tanpa
        # benar-benar memicu event pemilihan item. Solusi yang lebih
        # reliable: buka dulu dropdown-nya via safe click, lalu kirim
        # navigasi keyboard virtual-key via PostMessage (bukan send_keys,
        # yang gagal saat locked karena pakai SendInput).
        combo = frm.child_window(class_name="TwwDBComboBox", found_index=0)
        safe_set_focus(combo)
        ctl_sleep(0.1)

        WM_KEYDOWN = 0x0100
        WM_KEYUP   = 0x0101
        VK_HOME    = 0x24
        VK_DOWN    = 0x28
        VK_UP      = 0x26
        VK_RETURN  = 0x0D

        def send_vkey(hwnd, vk):
            ctypes.windll.user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0)
            ctl_sleep(0.01)
            ctypes.windll.user32.PostMessageW(hwnd, WM_KEYUP, vk, 0)

        def read_combo_text() -> str:
            try:
                return combo.window_text().strip()
            except Exception:
                return ""

        def open_dropdown() -> bool:
            dropdown_btn = frm.child_window(class_name="TBtnWinControl", found_index=0)
            if not click_control_message(dropdown_btn):
                log.warning("  Gagal klik dropdown Select Label (safe click gagal)")
                return False
            ctl_sleep(0.2)
            return True

        before_text = read_combo_text()
        log.info(f"  BROWN BOX - Isi Select Label sebelum diubah: '{before_text}'")

        target_hwnd = combo.handle
        selected_ok = False

        # Percobaan pertama: reset ke item pertama via Home, lalu Down
        # sebanyak BROWN_BOX_ITEM_INDEX kali (perilaku lama).
        # Percobaan berikutnya (kalau gagal): hitung SELISIH dari posisi
        # aktual (dibaca dari teks combo) ke posisi BROWN BOX, supaya tidak
        # asal tebak dari asumsi "combo pasti balik ke index 0".
        MAX_ATTEMPTS = 3
        for attempt in range(1, MAX_ATTEMPTS + 1):
            if not open_dropdown():
                return None

            current_text = read_combo_text()
            current_idx = BROWN_BOX_ITEMS_BY_TEXT.get(current_text.upper())

            if attempt == 1 or current_idx is None:
                # Tidak tahu posisi aktual (atau percobaan pertama) ->
                # pakai cara lama: Home dulu baru Down x N.
                send_vkey(target_hwnd, VK_HOME)
                ctl_sleep(0.1)
                steps, vk_dir = BROWN_BOX_ITEM_INDEX, VK_DOWN
            else:
                diff = BROWN_BOX_ITEM_INDEX - current_idx
                steps, vk_dir = abs(diff), (VK_DOWN if diff > 0 else VK_UP)

            for _ in range(steps):
                send_vkey(target_hwnd, vk_dir)
                ctl_sleep(0.025)
            ctl_sleep(0.1)
            send_vkey(target_hwnd, VK_RETURN)
            ctl_sleep(0.25)

            result_text = read_combo_text()
            log.info(f"  BROWN BOX - Percobaan {attempt}: Select Label -> '{result_text}'")

            if result_text.upper() == BROWN_BOX_LABEL_NAME.upper():
                selected_ok = True
                break
            log.warning(f"  BROWN BOX - Percobaan {attempt} salah pilih "
                        f"('{result_text}' != '{BROWN_BOX_LABEL_NAME}'), akan dicoba ulang")

        if not selected_ok:
            log.warning(f"  GAGAL set Select Label ke '{BROWN_BOX_LABEL_NAME}' "
                        f"setelah {MAX_ATTEMPTS}x percobaan, batalkan print Brown Box "
                        f"(supaya tidak mencetak label yang salah)")
            return None

        log.info(f"  OK: Select Label diset ke '{BROWN_BOX_LABEL_NAME}' (terverifikasi)")
    except Exception as e:
        log.warning(f"  Gagal set Select Label BROWN BOX: {e}")
        return None

    log.info("  BROWN BOX - Klik Print (TButton3)")
    try:
        if not click_control_message(frm.child_window(class_name="TButton", found_index=2)):
            log.warning("  Gagal klik TButton3 untuk BROWN BOX (safe click gagal)")
            return None
        ctl_sleep(0.5)
    except Exception as e:
        log.warning(f"  Gagal klik TButton3 untuk BROWN BOX: {e}")
        return None

    if not wait_win("TfrmPrint"):
        log.warning("  TfrmPrint tidak muncul untuk BROWN BOX")
        return None

    app = get_app()
    frm_print = app.window(class_name="TfrmPrint")
    if using_vdesk:
        move_window_to_automation_desktop(frm_print.handle)
        ctl_sleep(0.15)
    safe_set_focus(frm_print)
    ctl_sleep(0.25)

    return frm_print


def fill_and_print_brown_box(frm_print, box_code: str, using_vdesk: bool) -> bool:
    """
    Di TfrmPrint yang SUDAH TERBUKA: isi SKU tetap (Search By Plu Code) +
    kode box (Sku Code & Barcode Text), lalu klik Print. TfrmPrint TIDAK
    ditutup setelah ini, supaya kode box berikutnya bisa langsung diisi
    lagi tanpa balik ke menu awal. Pengisian TEdit tetap pakai
    set_edit_text (aman tanpa SendInput); hanya konfirmasi Enter dan klik
    tombol yang dialihkan ke jalur PostMessage.
    """
    ui_log(f"=== BROWN BOX | kode box: {box_code} ===")

    safe_set_focus(frm_print)
    ctl_sleep(0.15)

    edits = sorted(frm_print.children(class_name="TEdit"), key=lambda w: w.rectangle().top)
    log.info(f"  BROWN BOX - Ditemukan {len(edits)} TEdit di TfrmPrint")
    if len(edits) < 3:
        log.warning("  BROWN BOX - Jumlah TEdit di TfrmPrint tidak sesuai ekspektasi (< 3)")
        return False

    edit_plu   = edits[0]   # Search By Plu Code -> diisi SKU
    edit_sku   = edits[1]   # Sku Code (For Scanning) -> diisi kode box
    edit_bctxt = edits[2]   # Barcode Text -> diisi kode box

    WM_KEYDOWN = 0x0100
    WM_KEYUP   = 0x0101
    VK_RETURN  = 0x0D

    def send_enter(hwnd):
        ctypes.windll.user32.PostMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0)
        ctl_sleep(0.01)
        ctypes.windll.user32.PostMessageW(hwnd, WM_KEYUP, VK_RETURN, 0)

    log.info(f"  BROWN BOX - Isi SKU (Search By Plu Code): {BROWN_BOX_SKU}")
    try:
        safe_set_focus(edit_plu)
        ctl_sleep(0.1)
        edit_plu.set_edit_text(BROWN_BOX_SKU)
        send_enter(edit_plu.handle)
        ctl_sleep(0.25)
    except Exception as e:
        log.warning(f"  Gagal isi SKU (Search By Plu Code): {e}")
        return False

    log.info(f"  BROWN BOX - Isi kode box (Sku Code & Barcode Text): {box_code}")
    try:
        safe_set_focus(edit_sku)
        ctl_sleep(0.1)
        edit_sku.set_edit_text(box_code)
        ctl_sleep(0.1)

        safe_set_focus(edit_bctxt)
        ctl_sleep(0.1)
        edit_bctxt.set_edit_text(box_code)
        ctl_sleep(0.1)
    except Exception as e:
        log.warning(f"  Gagal isi Sku Code/Barcode Text (kode box): {e}")
        return False

    log.info("  BROWN BOX - Klik Print (TcxButton12)")
    try:
        print_btn = frm_print.child_window(class_name="TcxButton", found_index=11)
        if not click_control_message(print_btn):
            log.warning("  Gagal klik TcxButton12 (Print) (safe click gagal)")
            return False
        ctl_sleep(0.25)
    except Exception as e:
        log.warning(f"  Gagal klik TcxButton12 (Print): {e}")
        return False

    popup = wait_win("TMessageForm", timeout=0.5)
    if popup:
        try:
            app = get_app()
            dlg = app.window(class_name="TMessageForm")
            if using_vdesk:
                move_window_to_automation_desktop(dlg.handle)
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            ctl_sleep(0.25)
            log.info("  OK: Popup konfirmasi BROWN BOX ditutup")
        except Exception as e:
            log.info(f"  Popup konfirmasi BROWN BOX error: {e}, lanjut")
    else:
        log.info("  Tidak ada popup konfirmasi BROWN BOX, lanjut")

    ui_log(f"=== BROWN BOX SELESAI | {box_code} ===")
    return True


def print_brown_box_labels(box_codes: list[str]) -> tuple[list[str], list[str]]:
    using_vdesk = create_automation_desktop()
    success_codes, failed_codes = [], []
    processed = 0

    try:
        frm_print = open_brown_box_print_form(using_vdesk)
        if frm_print is None:
            return [], list(box_codes)

        for box_code in box_codes:
            if fill_and_print_brown_box(frm_print, box_code, using_vdesk):
                success_codes.append(box_code)
            else:
                failed_codes.append(box_code)
            processed += 1

        return success_codes, failed_codes

    except StopRequested:
        ui_log("BROWN BOX dihentikan manual oleh pengguna.", level=logging.WARNING)
        failed_codes.extend(box_codes[processed:])
        if STOP_CLOSES_EBARCODE:
            with control.cleanup():
                try:
                    kill_ebarcode()
                except Exception:
                    pass
        return success_codes, failed_codes

    finally:
        if using_vdesk:
            destroy_automation_desktop()


def open_gondola_print_form(using_vdesk: bool):
    """
    Buka EBarcode (kalau belum jalan) -> di TfrmBarcode pilih 'GONDOLA' di
    Select Label -> klik Print (TButton3) -> tunggu TfrmPrint muncul.
    Return: frm_print (pywinauto window) kalau berhasil, None kalau gagal.
    Sama persis dengan open_brown_box_print_form(), hanya beda label target.
    """
    ui_log("  GONDOLA - Buka EBarcode")

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    leftover_print = Desktop(backend="win32").windows(class_name="TfrmPrint")
    if leftover or leftover_print:
        log.warning("  GONDOLA - EBarcode/TfrmPrint masih terbuka dari sesi "
                    "sebelumnya, tutup dulu supaya tidak salah pilih menu...")
        kill_ebarcode()
        ctl_sleep(0.75)

    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        ctl_sleep(1.5)

    if not wait_win("TfrmBarcode", timeout=TIMEOUT):
        log.warning("  TfrmBarcode tidak muncul, lewati print GONDOLA")
        return None

    app = get_app()
    frm = app.window(class_name="TfrmBarcode")
    if using_vdesk:
        move_window_to_automation_desktop(frm.handle)
        ctl_sleep(0.15)
    safe_set_focus(frm)
    ctl_sleep(0.25)

    log.info(f"  GONDOLA - Pilih label '{GONDOLA_LABEL_NAME}' di Select Label")
    try:
        combo = frm.child_window(class_name="TwwDBComboBox", found_index=0)
        safe_set_focus(combo)
        ctl_sleep(0.1)

        WM_KEYDOWN = 0x0100
        WM_KEYUP   = 0x0101
        VK_HOME    = 0x24
        VK_DOWN    = 0x28
        VK_UP      = 0x26
        VK_RETURN  = 0x0D

        def send_vkey(hwnd, vk):
            ctypes.windll.user32.PostMessageW(hwnd, WM_KEYDOWN, vk, 0)
            ctl_sleep(0.01)
            ctypes.windll.user32.PostMessageW(hwnd, WM_KEYUP, vk, 0)

        def read_combo_text() -> str:
            try:
                return combo.window_text().strip()
            except Exception:
                return ""

        def open_dropdown() -> bool:
            dropdown_btn = frm.child_window(class_name="TBtnWinControl", found_index=0)
            if not click_control_message(dropdown_btn):
                log.warning("  Gagal klik dropdown Select Label (safe click gagal)")
                return False
            ctl_sleep(0.2)
            return True

        before_text = read_combo_text()
        log.info(f"  GONDOLA - Isi Select Label sebelum diubah: '{before_text}'")

        target_hwnd = combo.handle
        selected_ok = False

        MAX_ATTEMPTS = 3
        for attempt in range(1, MAX_ATTEMPTS + 1):
            if not open_dropdown():
                return None

            current_text = read_combo_text()
            current_idx = BROWN_BOX_ITEMS_BY_TEXT.get(current_text.upper())

            if attempt == 1 or current_idx is None:
                send_vkey(target_hwnd, VK_HOME)
                ctl_sleep(0.1)
                steps, vk_dir = GONDOLA_ITEM_INDEX, VK_DOWN
            else:
                diff = GONDOLA_ITEM_INDEX - current_idx
                steps, vk_dir = abs(diff), (VK_DOWN if diff > 0 else VK_UP)

            for _ in range(steps):
                send_vkey(target_hwnd, vk_dir)
                ctl_sleep(0.025)
            ctl_sleep(0.1)
            send_vkey(target_hwnd, VK_RETURN)
            ctl_sleep(0.25)

            result_text = read_combo_text()
            log.info(f"  GONDOLA - Percobaan {attempt}: Select Label -> '{result_text}'")

            if result_text.upper() == GONDOLA_LABEL_NAME.upper():
                selected_ok = True
                break
            log.warning(f"  GONDOLA - Percobaan {attempt} salah pilih "
                        f"('{result_text}' != '{GONDOLA_LABEL_NAME}'), akan dicoba ulang")

        if not selected_ok:
            log.warning(f"  GAGAL set Select Label ke '{GONDOLA_LABEL_NAME}' "
                        f"setelah {MAX_ATTEMPTS}x percobaan, batalkan print Gondola "
                        f"(supaya tidak mencetak label yang salah)")
            return None

        log.info(f"  OK: Select Label diset ke '{GONDOLA_LABEL_NAME}' (terverifikasi)")
    except Exception as e:
        log.warning(f"  Gagal set Select Label GONDOLA: {e}")
        return None

    log.info("  GONDOLA - Klik Print (TButton3)")
    try:
        if not click_control_message(frm.child_window(class_name="TButton", found_index=2)):
            log.warning("  Gagal klik TButton3 untuk GONDOLA (safe click gagal)")
            return None
        ctl_sleep(0.5)
    except Exception as e:
        log.warning(f"  Gagal klik TButton3 untuk GONDOLA: {e}")
        return None

    if not wait_win("TfrmPrint"):
        log.warning("  TfrmPrint tidak muncul untuk GONDOLA")
        return None

    app = get_app()
    frm_print = app.window(class_name="TfrmPrint")
    if using_vdesk:
        move_window_to_automation_desktop(frm_print.handle)
        ctl_sleep(0.15)
    safe_set_focus(frm_print)
    ctl_sleep(0.25)

    return frm_print


def fill_and_print_gondola(frm_print, box_code: str, using_vdesk: bool) -> bool:
    """
    Di TfrmPrint yang SUDAH TERBUKA: isi SKU tetap (Search By Plu Code) +
    kode (Sku Code & Barcode Text), lalu klik Print. TfrmPrint TIDAK
    ditutup setelah ini, supaya kode berikutnya bisa langsung diisi lagi
    tanpa balik ke menu awal. Sama persis dengan fill_and_print_brown_box(),
    hanya beda SKU/label target dan teks log.
    """
    ui_log(f"=== GONDOLA | kode: {box_code} ===")

    safe_set_focus(frm_print)
    ctl_sleep(0.15)

    edits = sorted(frm_print.children(class_name="TEdit"), key=lambda w: w.rectangle().top)
    log.info(f"  GONDOLA - Ditemukan {len(edits)} TEdit di TfrmPrint")
    if len(edits) < 3:
        log.warning("  GONDOLA - Jumlah TEdit di TfrmPrint tidak sesuai ekspektasi (< 3)")
        return False

    edit_plu   = edits[0]   # Search By Plu Code -> diisi SKU
    edit_sku   = edits[1]   # Sku Code (For Scanning) -> diisi kode
    edit_bctxt = edits[2]   # Barcode Text -> diisi kode

    WM_KEYDOWN = 0x0100
    WM_KEYUP   = 0x0101
    VK_RETURN  = 0x0D

    def send_enter(hwnd):
        ctypes.windll.user32.PostMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0)
        ctl_sleep(0.01)
        ctypes.windll.user32.PostMessageW(hwnd, WM_KEYUP, VK_RETURN, 0)

    log.info(f"  GONDOLA - Isi SKU (Search By Plu Code): {GONDOLA_SKU}")
    try:
        safe_set_focus(edit_plu)
        ctl_sleep(0.1)
        edit_plu.set_edit_text(GONDOLA_SKU)
        send_enter(edit_plu.handle)
        ctl_sleep(0.25)
    except Exception as e:
        log.warning(f"  Gagal isi SKU (Search By Plu Code): {e}")
        return False

    log.info(f"  GONDOLA - Isi kode (Sku Code & Barcode Text): {box_code}")
    try:
        safe_set_focus(edit_sku)
        ctl_sleep(0.1)
        edit_sku.set_edit_text(box_code)
        ctl_sleep(0.1)

        safe_set_focus(edit_bctxt)
        ctl_sleep(0.1)
        edit_bctxt.set_edit_text(box_code)
        ctl_sleep(0.1)
    except Exception as e:
        log.warning(f"  Gagal isi Sku Code/Barcode Text (kode): {e}")
        return False

    log.info("  GONDOLA - Klik Print (TcxButton12)")
    try:
        print_btn = frm_print.child_window(class_name="TcxButton", found_index=11)
        if not click_control_message(print_btn):
            log.warning("  Gagal klik TcxButton12 (Print) (safe click gagal)")
            return False
        ctl_sleep(0.25)
    except Exception as e:
        log.warning(f"  Gagal klik TcxButton12 (Print): {e}")
        return False

    popup = wait_win("TMessageForm", timeout=0.5)
    if popup:
        try:
            app = get_app()
            dlg = app.window(class_name="TMessageForm")
            if using_vdesk:
                move_window_to_automation_desktop(dlg.handle)
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            ctl_sleep(0.25)
            log.info("  OK: Popup konfirmasi GONDOLA ditutup")
        except Exception as e:
            log.info(f"  Popup konfirmasi GONDOLA error: {e}, lanjut")
    else:
        log.info("  Tidak ada popup konfirmasi GONDOLA, lanjut")

    ui_log(f"=== GONDOLA SELESAI | {box_code} ===")
    return True


def print_gondola_labels(box_codes: list[str]) -> tuple[list[str], list[str]]:
    using_vdesk = create_automation_desktop()
    success_codes, failed_codes = [], []
    processed = 0

    try:
        frm_print = open_gondola_print_form(using_vdesk)
        if frm_print is None:
            return [], list(box_codes)

        for box_code in box_codes:
            if fill_and_print_gondola(frm_print, box_code, using_vdesk):
                success_codes.append(box_code)
            else:
                failed_codes.append(box_code)
            processed += 1

        return success_codes, failed_codes

    except StopRequested:
        ui_log("GONDOLA dihentikan manual oleh pengguna.", level=logging.WARNING)
        failed_codes.extend(box_codes[processed:])
        if STOP_CLOSES_EBARCODE:
            with control.cleanup():
                try:
                    kill_ebarcode()
                except Exception:
                    pass
        return success_codes, failed_codes

    finally:
        if using_vdesk:
            destroy_automation_desktop()


# ============================================================
# BAGIAN 3: ALUR GABUNGAN PER TOMBOL
# ============================================================

def _read_nonempty_lines(path) -> list:
    try:
        txt = Path(path).read_text(encoding="utf-8", errors="ignore")
    except FileNotFoundError:
        return []
    return [l.strip() for l in txt.splitlines() if l.strip()]


def process_kind(kind: str, sync_result: dict = None) -> tuple[bool, str]:
    """Menjalankan otomasi EBarcode (buka aplikasi -> cetak) untuk satu kind
    ('IN' atau 'OUT').

    - Data besar dipecah per BATCH_SIZE baris supaya tiap putaran EBarcode ringan.
    - Jika data berasal dari Sync (sync_result punya item + state), state di-commit
      PER BATCH hanya setelah batch itu sukses dicetak. Batch yang gagal tidak
      di-commit, jadi otomatis muncul lagi di Sync berikutnya (tidak hilang).
    - Jika sync_result tidak punya item (Reprint / Recon), baris dibaca dari file
      yang ada di disk dan diproses tanpa commit state.
    """
    if kind == "IN":
        source_file, label_file = Path(FILE1), Path(FILE2)
        label_search, items_key = LABEL_SEARCH_IN, "in_items"
    elif kind == "OUT":
        source_file, label_file = Path(FILE3), Path(FILE4)
        label_search, items_key = LABEL_SEARCH_OUT, "out_items"
    else:
        return False, f"Kind tidak dikenal: {kind}"

    sync_result = sync_result or {}
    items = sync_result.get(items_key) or None
    state = sync_result.get("state")

    if items is None:
        box_lines = _read_nonempty_lines(source_file)
        label_lines = _read_nonempty_lines(label_file)
        if not box_lines:
            return False, f"Tidak ada data untuk diproses.\nFile kosong/tidak ada:\n{source_file}"
        if len(label_lines) == len(box_lines):
            items = [{"box": b, "label": l, "keys": [], "box_code": None}
                     for b, l in zip(box_lines, label_lines)]
        else:
            log.warning(f"  Jumlah baris '{source_file.name}' ({len(box_lines)}) != "
                        f"'{label_file.name}' ({len(label_lines)}); diproses sebagai satu batch.")

    # Susun batch. None = jalankan apa adanya dari file yang ada (tidak bisa dipecah).
    if items:
        batches = [items[i:i + BATCH_SIZE] for i in range(0, len(items), BATCH_SIZE)]
    else:
        batches = [None]

    total_batches = len(batches)
    done_items, failed_at, stopped = [], None, False

    for idx, batch in enumerate(batches, start=1):
        if batch is not None:
            write_txt(str(source_file), [it["box"] for it in batch])
            write_txt(str(label_file), [it["label"] for it in batch])
            ui_log(f"{kind}: batch {idx}/{total_batches} ({len(batch)} baris)")

        if control.stop_requested:          # STOP ditekan di sela batch / sebelum mulai
            failed_at, stopped = idx, True
            break
        if not run_print_flow(source_file, label_search):
            failed_at = idx
            stopped = control.stop_requested
            break

        if batch is not None:
            done_items.extend(batch)
            keys = [k for it in batch for k in it.get("keys", [])]
            if state is not None and keys:
                try:
                    if kind == "IN":
                        state.commit_in(keys)
                    else:
                        state.commit_out(keys)
                except Exception as e:
                    log.error(f"  Gagal menyimpan state batch {idx}: {e}")
                    failed_at = idx
                    break

    # Kembalikan isi file: sisa yang belum selesai (kalau gagal) atau seluruh batch (kalau sukses),
    # supaya tombol Reprint tetap konsisten.
    if items:
        remaining = items[len(done_items):] if failed_at else items
        try:
            write_txt(str(source_file), [it["box"] for it in remaining])
            write_txt(str(label_file), [it["label"] for it in remaining])
        except Exception as e:
            log.warning(f"  Gagal memulihkan isi file {kind}: {e}")

    brown_box_summary = ""
    if kind == "IN":
        if items and done_items and done_items[0].get("box_code") is not None:
            box_codes = sorted({it["box_code"] for it in done_items if it.get("box_code")})
        elif failed_at is None:
            box_codes = sync_result.get("new_box_codes") or []
        else:
            box_codes = []
        if box_codes and stopped:
            # Dihentikan manual: jangan cetak Brown Box, tapi simpan daftarnya
            # (batch yang sudah ter-commit tidak akan muncul lagi di Sync berikutnya).
            pending = os.path.join(DIR_PROSES, "BROWN BOX PENDING.txt")
            try:
                write_txt(pending, box_codes)
            except Exception as e:
                log.warning(f"Gagal menulis {pending}: {e}")
            brown_box_summary = (
                f"\n\n\u25CF Brown Box BELUM dicetak ({len(box_codes)} kode). Daftar disimpan di:\n"
                f"{pending}\nSalin isinya ke tombol 'Print Brown Box (Custom)' bila perlu."
            )
        elif box_codes:
            ui_log(f"BROWN BOX - {len(box_codes)} kode box baru akan dicetak: {box_codes}")
            success_codes, failed_codes = print_brown_box_labels(box_codes)
            brown_box_summary = f"\n\n\u25CF Brown Box tercetak : {', '.join(success_codes) or '-'}"
            if failed_codes:
                brown_box_summary += f"\n\u25CF Brown Box gagal/dihentikan: {', '.join(failed_codes)}"
            if control.stop_requested and failed_codes:
                stopped = True
        else:
            ui_log("BROWN BOX - Tidak ada kode box baru, menu Brown Box dilewati")

    if failed_at is not None:
        sisa = (len(items) - len(done_items)) if items else "?"
        judul = "DIHENTIKAN manual" if stopped else "BERHENTI"
        return False, (
            f"Proses {kind} {judul} di batch {failed_at}/{total_batches}. "
            f"Selesai: {len(done_items)} baris, sisa: {sisa} baris.\n"
            f"Sisa data TIDAK hilang: akan muncul lagi saat Sync berikutnya "
            f"(atau tekan Reprint). Lihat log.txt untuk detail."
            f"{brown_box_summary}"
        )

    if stopped:   # dihentikan saat cetak Brown Box (semua batch utama sudah selesai)
        return False, (f"Proses {kind} selesai, tetapi cetak Brown Box DIHENTIKAN manual."
                       f"{brown_box_summary}")

    batch_info = f" ({total_batches} batch)" if total_batches > 1 else ""
    return True, (
        f"Proses {kind} berhasil{batch_info}.\n\n"
        f"\u25CF Sumber data : {source_file}\n"
        f"\u25CF Label print : {label_search}"
        f"{brown_box_summary}"
    )


def run_flow_auto() -> tuple[bool, str]:
    """[UPDATE KS] - Menu tunggal (gabungan IN & OUT).
    Ambil data dari Google Sheets sekali, lalu deteksi otomatis:
      - Ada data IN  -> proses IN  (EXCEL IN BOX.txt  & +PRINT IN LABEL.txt)
      - Ada data OUT -> proses OUT (EXCEL OUT BOX.txt & -PRINT OUT LABEL.txt)
    Kalau dua-duanya ada, dua-duanya diproses berurutan (IN dulu, lalu OUT).
    """
    if not HAS_PYWINAUTO:
        return False, "pywinauto tidak terinstall. Jalankan: pip install pywinauto"

    sync_result = sync_from_sheet()
    if not sync_result["ok"]:
        return False, f"Gagal ambil data dari spreadsheet:\n{sync_result['error']}"

    if sync_result.get("baseline") is not None:
        return True, (
            f"Pemakaian pertama: {sync_result['baseline']} data yang sudah ada dicatat sebagai "
            f"baseline (TIDAK dicetak).\n\nMulai sekarang hanya data BARU / yang keluar "
            f"yang akan diproses saat Sync."
        )

    found_new = sync_result.get("found_new")
    found_out = sync_result.get("found_out")

    if not found_new and not found_out:
        return True, "Sinkronisasi selesai. Tidak ada data IN maupun OUT baru untuk diproses."

    messages = []
    overall_ok = True

    if found_new:
        ui_log("Terdeteksi data IN baru -> menjalankan proses KS IN...")
        ok, msg = process_kind("IN", sync_result)
        messages.append(msg)
        overall_ok = overall_ok and ok

    if found_out and not control.stop_requested:
        ui_log("Terdeteksi data OUT baru -> menjalankan proses KS OUT...")
        ok, msg = process_kind("OUT", sync_result)
        messages.append(msg)
        overall_ok = overall_ok and ok

    return overall_ok, "\n\n".join(messages)


def run_flow_recon() -> tuple[bool, str]:
    """[SYNC RECON] - Rekonsiliasi berdasarkan status per-baris pada kolom BQ.
    Membaca SKU dari kolom BM ke bawah, mengecek status dari kolom BQ ke
    bawah, lalu menulis EXCEL IN/OUT BOX.txt & +PRINT IN/OUT LABEL.txt sesuai
    status baris tersebut, dan menjalankan otomasi EBarcode berdasarkan hasilnya.
    """
    if not HAS_PYWINAUTO:
        return False, "pywinauto tidak terinstall. Jalankan: pip install pywinauto"

    recon_result = reconcile_ks_master()
    if not recon_result["ok"]:
        return False, f"Gagal rekonsiliasi data:\n{recon_result['error']}"

    found_in  = recon_result.get("found_in")
    found_out = recon_result.get("found_out")

    if not found_in and not found_out:
        return True, "Rekonsiliasi selesai. Tidak ada selisih 'Belum Excel In/Out' untuk diproses."

    messages = []
    overall_ok = True

    if found_in:
        ui_log("KS RECON: terdeteksi SKU 'Belum Excel In' -> menjalankan proses IN...")
        ok, msg = process_kind("IN", recon_result)
        messages.append(msg)
        overall_ok = overall_ok and ok

    if found_out and not control.stop_requested:
        ui_log("KS RECON: terdeteksi SKU 'Belum Excel Out' -> menjalankan proses OUT...")
        ok, msg = process_kind("OUT", recon_result)
        messages.append(msg)
        overall_ok = overall_ok and ok

    return overall_ok, "\n\n".join(messages)


def run_flow_repeat() -> tuple[bool, str]:
    """[ULANGI PROSES EBARCODE] - Untuk kasus user lupa apakah proses EBarcode
    sudah dijalankan atau belum. TIDAK mengambil data baru dari Google Sheets,
    langsung baca file yang sudah ada (EXCEL IN BOX.txt / EXCEL OUT BOX.txt)
    dan ulangi dari langkah buka aplikasi EBarcode.
    """
    if not HAS_PYWINAUTO:
        return False, "pywinauto tidak terinstall. Jalankan: pip install pywinauto"

    file_in  = Path(FILE1)
    file_out = Path(FILE3)
    has_in  = file_in.exists()  and bool(file_in.read_text(encoding="utf-8", errors="ignore").strip())
    has_out = file_out.exists() and bool(file_out.read_text(encoding="utf-8", errors="ignore").strip())

    if not has_in and not has_out:
        return False, (
            "Tidak ditemukan data tersisa untuk diulang di:\n"
            f"{file_in}\n{file_out}\n\n"
            "Gunakan tombol UPDATE KS untuk mengambil data baru dari spreadsheet."
        )

    ui_log("===== ULANGI PROSES EBARCODE (tanpa ambil data Google Sheets) =====")

    messages = []
    overall_ok = True

    if has_in:
        ui_log("Ditemukan sisa data IN -> mengulang otomasi EBarcode untuk IN...")
        ok, msg = process_kind("IN", None)
        messages.append(msg)
        overall_ok = overall_ok and ok

    if has_out and not control.stop_requested:
        ui_log("Ditemukan sisa data OUT -> mengulang otomasi EBarcode untuk OUT...")
        ok, msg = process_kind("OUT", None)
        messages.append(msg)
        overall_ok = overall_ok and ok

    if has_in:
        messages.append(
            "\u26A0 Brown Box TIDAK ikut dicetak ulang otomatis pada mode ini "
            "(kode box baru hanya diketahui saat sinkronisasi spreadsheet). "
            "Cetak manual dari menu Brown Box bila perlu."
        )

    return overall_ok, "\n\n".join(messages)


# ============================================================
# BAGIAN 4: GUI
# ============================================================

if HAS_TK:
    class TextHandler(logging.Handler):
        def __init__(self, text_widget):
            super().__init__()
            self.text_widget = text_widget

        def emit(self, record):
            if not getattr(record, "show_ui", False):
                return

            msg = self.format(record)

            def append():
                self.text_widget.configure(state="normal")
                self.text_widget.delete("1.0", "end")
                self.text_widget.insert("end", msg)
                self.text_widget.configure(state="disabled")

            try:
                self.text_widget.after(0, append)
            except Exception:
                pass


    C_BG          = "#EAF2FA"
    C_HEADER_BG   = "#AECDEA"
    C_HEADER_FG   = "#1F3A5C"
    C_PRIMARY     = "#7FB0DE"
    C_PRIMARY_HOV = "#6BA0D2"
    C_SECONDARY   = "#93B8D8"
    C_SECONDARY_HOV = "#7FA6C9"
    C_TEXT        = "#2A4A68"
    C_MUTED       = "#71879E"
    C_CARD_BG     = "#FFFFFF"
    C_BORDER      = "#DCE9F6"
    C_LOG_BG      = "#F3F8FD"
    C_LOG_FG      = "#33526E"
    C_STATUS      = "#5686B5"
    C_GEAR_BG     = "#93B8D8"
    C_GEAR_HOV    = "#7FA6C9"
    C_RECON       = "#9BC6AC"
    C_RECON_HOV   = "#87B999"
    C_STOP        = "#E8A9A9"
    C_STOP_HOV    = "#DE9494"
    C_PAUSE       = "#F2D49B"
    C_PAUSE_HOV   = "#E8C77F"


    class StoreSelectDialog(tk.Toplevel):
        """Dialog pilih toko aktif - hanya muncul SEKALI di awal (first run),
        SELALU muncul dan wajib dipilih manual walaupun daftar toko cuma satu
        (toko default). User bisa langsung pilih toko yang ada, atau menambah
        toko baru dulu lewat tombol '+ Tambah Toko Baru'.
        Pilihan disimpan permanen (mark_store_selected), jadi buka aplikasi
        berikutnya langsung otomatis pakai toko yang sama tanpa tanya lagi.
        Satu-satunya cara berubah setelah itu adalah lewat menu Pengaturan."""

        def __init__(self, parent):
            super().__init__(parent)
            self.title("Pilih Toko")
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.protocol("WM_DELETE_WINDOW", lambda: None)  # wajib pilih, tidak boleh ditutup

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            self.wrap = tk.Frame(self, bg=C_CARD_BG, padx=26, pady=20)
            self.wrap.pack(fill="both", expand=True)

            tk.Label(self.wrap, text="Pilih toko aktif", font=("Segoe UI", 12, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w")
            tk.Label(self.wrap,
                     text="Wajib dipilih manual di awal (sekali saja).\n"
                          "Setelah itu pilihan diingat otomatis setiap buka aplikasi.",
                     font=("Segoe UI", 8), fg=C_MUTED, bg=C_CARD_BG, justify="left"
                     ).pack(anchor="w", pady=(2, 12))

            self.var_code = tk.StringVar(value="")  # kosong: tidak ada toko aktif default

            self.radio_frame = tk.Frame(self.wrap, bg=C_CARD_BG)
            self.radio_frame.pack(fill="x")
            self._render_radios()

            btn_add = tk.Button(
                self.wrap, text="+ Tambah Toko Baru", font=("Segoe UI", 9, "bold"),
                bg="#EAF2FD", fg=C_PRIMARY, relief="flat", padx=12, pady=6,
                activebackground="#DCEAFB", activeforeground=C_PRIMARY,
                cursor="hand2", bd=0, command=self._add_store
            )
            btn_add.pack(fill="x", pady=(10, 0))

            self.btn_ok = tk.Button(
                self.wrap, text="Pilih & Mulai", font=("Segoe UI", 10, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=16, pady=8,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._confirm
            )
            self.btn_ok.pack(fill="x", pady=(14, 0))

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")
            self.deiconify()
            self.lift()
            self.focus_force()
            try:
                self.grab_set()
            except Exception as e:
                log.error(f"Gagal grab_set() pada dialog pilih toko: {e}")

        def _render_radios(self):
            for child in self.radio_frame.winfo_children():
                child.destroy()
            if not STORES:
                tk.Label(self.radio_frame, text="Belum ada toko. Tambah dulu di bawah.",
                          font=("Segoe UI", 9, "italic"), fg=C_MUTED, bg=C_CARD_BG
                          ).pack(anchor="w", pady=2)
                return
            for s in STORES:
                tk.Radiobutton(
                    self.radio_frame, text=s["code"], variable=self.var_code, value=s["code"],
                    font=("Segoe UI", 10), fg=C_TEXT, bg=C_CARD_BG,
                    selectcolor="#FFFFFF", activebackground=C_CARD_BG, anchor="w"
                ).pack(fill="x", pady=2)

        def _add_store(self):
            def on_submit(code, sid, gid):
                upsert_store(code, sid, gid, activate=False)
                self._render_radios()
                self.var_code.set(code)  # langsung tandai toko baru sebagai pilihan
                ui_log(f"Toko baru ditambahkan saat pilih toko awal: {code}")
            StoreFormDialog(self, "Tambah Toko", on_submit=on_submit)

        def _confirm(self):
            code = self.var_code.get().strip()
            if not code:
                messagebox.showwarning(
                    "Pilih Toko",
                    "Pilih salah satu toko dulu, atau tambah toko baru.",
                    parent=self
                )
                return
            set_active_store(code)
            mark_store_selected()
            self.destroy()


    class StoreFormDialog(tk.Toplevel):
        def __init__(self, parent, title: str, initial: dict = None, on_submit=None):
            super().__init__(parent)
            self.on_submit = on_submit
            self.initial = initial or {}
            self.title(title)
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            wrap = tk.Frame(self, bg=C_CARD_BG, padx=30, pady=24)
            wrap.pack(fill="both", expand=True)

            tk.Label(wrap, text=title, font=("Segoe UI", 13, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w", pady=(0, 14))

            fields = [
                ("Kode Toko",      self.initial.get("code", ""),           "Contoh: CGAR"),
                ("Spreadsheet ID", self.initial.get("spreadsheet_id", ""), "ID panjang pada URL Google Sheets"),
                ("GID",            self.initial.get("gid", ""),            "ID sheet/tab pada spreadsheet"),
            ]
            self.entries = {}
            for label, value, hint in fields:
                block = tk.Frame(wrap, bg=C_CARD_BG)
                block.pack(fill="x", pady=(0, 12))

                tk.Label(block, text=label, font=("Segoe UI", 9, "bold"),
                         fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w")

                holder = tk.Frame(block, bg=C_BORDER)
                holder.pack(fill="x", pady=(4, 0))
                ent = tk.Entry(holder, font=("Segoe UI", 10), width=42,
                                relief="flat", bd=0, fg=C_TEXT, bg="#FBFDFF",
                                insertbackground=C_TEXT)
                ent.insert(0, value)
                ent.pack(fill="x", padx=1, pady=1, ipady=7, ipadx=8)
                ent.bind("<FocusIn>", lambda e, h=holder: h.config(bg=C_PRIMARY))
                ent.bind("<FocusOut>", lambda e, h=holder: h.config(bg=C_BORDER))

                tk.Label(block, text=hint, font=("Segoe UI", 8),
                         fg=C_MUTED, bg=C_CARD_BG).pack(anchor="w", pady=(3, 0))

                self.entries[label] = ent

            btn_row = tk.Frame(wrap, bg=C_CARD_BG)
            btn_row.pack(fill="x", pady=(10, 0))

            btn_save = tk.Button(
                btn_row, text="Simpan", font=("Segoe UI", 10, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=18, pady=9,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._submit
            )
            btn_save.pack(side="right")
            bind_hover(btn_save, C_PRIMARY, C_PRIMARY_HOV)

            btn_cancel = tk.Button(
                btn_row, text="Batal", font=("Segoe UI", 10, "bold"),
                bg="#F1F5FB", fg=C_MUTED, relief="flat", padx=16, pady=9,
                activebackground="#E3ECF7", cursor="hand2", bd=0,
                command=self.destroy
            )
            btn_cancel.pack(side="right", padx=(0, 10))
            bind_hover(btn_cancel, "#F1F5FB", "#E3ECF7")

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

        def _submit(self):
            code = self.entries["Kode Toko"].get().strip()
            sid  = self.entries["Spreadsheet ID"].get().strip()
            gid  = self.entries["GID"].get().strip()

            if not code or not sid or not gid:
                messagebox.showwarning("Toko", "Semua field wajib diisi.", parent=self)
                return

            if self.on_submit:
                try:
                    self.on_submit(code, sid, gid)
                except ValueError as e:
                    messagebox.showwarning("Toko", str(e), parent=self)
                    return

            self.destroy()


    def bind_hover(widget, normal_bg, hover_bg):
        widget.bind("<Enter>", lambda e: widget.config(bg=hover_bg))
        widget.bind("<Leave>", lambda e: widget.config(bg=normal_bg))


    class SettingsDialog(tk.Toplevel):
        def __init__(self, parent, on_saved=None):
            super().__init__(parent)
            self.on_saved = on_saved
            self.title("Pengaturan")
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            wrap = tk.Frame(self, bg=C_CARD_BG, padx=16, pady=14)
            wrap.pack(fill="both", expand=True)

            title_row = tk.Frame(wrap, bg=C_CARD_BG)
            title_row.pack(fill="x", pady=(0, 2))
            tk.Label(title_row, text="\u2699", font=("Segoe UI", 11),
                     fg=C_PRIMARY, bg=C_CARD_BG).pack(side="left", padx=(0, 5))
            tk.Label(title_row, text="Pengaturan Toko", font=("Segoe UI", 10, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(side="left")

            tk.Label(wrap, text="Kelola daftar toko & sumber data spreadsheet.",
                     font=("Segoe UI", 8), fg=C_MUTED, bg=C_CARD_BG
                     ).pack(fill="x", pady=(0, 2))

            tk.Frame(wrap, bg=C_BORDER, height=1).pack(fill="x", pady=(6, 8))

            style = ttk.Style(self)
            try:
                style.theme_use("clam")
            except Exception:
                pass
            style.configure("Store.Treeview", font=("Segoe UI", 8), rowheight=22,
                             background="#FFFFFF", fieldbackground="#FFFFFF",
                             foreground=C_TEXT, borderwidth=0)
            style.configure("Store.Treeview.Heading", font=("Segoe UI", 8, "bold"),
                             background=C_BORDER, foreground=C_TEXT, relief="flat")
            style.map("Store.Treeview",
                      background=[("selected", C_PRIMARY)],
                      foreground=[("selected", "#FFFFFF")])

            table_frame = tk.Frame(wrap, bg=C_CARD_BG)
            table_frame.pack(fill="both", expand=True)

            columns = ("code", "spreadsheet_id", "gid", "status")
            self.tree = ttk.Treeview(
                table_frame, columns=columns, show="headings", height=4,
                style="Store.Treeview", selectmode="browse"
            )
            self.tree.heading("code", text="Kode Toko")
            self.tree.heading("spreadsheet_id", text="Spreadsheet ID")
            self.tree.heading("gid", text="GID")
            self.tree.heading("status", text="Status")
            self.tree.column("code", width=70, anchor="w")
            self.tree.column("spreadsheet_id", width=180, anchor="w")
            self.tree.column("gid", width=70, anchor="center")
            self.tree.column("status", width=55, anchor="center")
            self.tree.pack(side="left", fill="both", expand=True)

            scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
            scrollbar.pack(side="right", fill="y")
            self.tree.configure(yscrollcommand=scrollbar.set)

            self._reload_table()

            action_row = tk.Frame(wrap, bg=C_CARD_BG)
            action_row.pack(fill="x", pady=(6, 2))

            def make_btn(parent_, text, cmd, bg, fg, hover):
                b = tk.Button(parent_, text=text, font=("Segoe UI", 8, "bold"),
                              bg=bg, fg=fg, relief="flat", padx=8, pady=4,
                              activebackground=hover, activeforeground=fg,
                              cursor="hand2", bd=0, command=cmd)
                bind_hover(b, bg, hover)
                return b

            make_btn(action_row, "+ Tambah", self._add_store,
                      C_PRIMARY, "#FFFFFF", C_PRIMARY_HOV).pack(side="left")
            make_btn(action_row, "Edit", self._edit_store,
                      "#EAF2FD", C_PRIMARY, "#DCEAFB").pack(side="left", padx=(6, 0))
            make_btn(action_row, "Aktifkan", self._activate_store,
                      "#EAF2FD", C_PRIMARY, "#DCEAFB").pack(side="left", padx=(6, 0))
            make_btn(action_row, "Hapus", self._delete_store,
                      "#FDEEEE", "#D32F2F", "#FADCDC").pack(side="left", padx=(6, 0))

            tk.Frame(wrap, bg=C_BORDER, height=1).pack(fill="x", pady=(10, 10))

            path_title_row = tk.Frame(wrap, bg=C_CARD_BG)
            path_title_row.pack(fill="x", pady=(0, 2))
            tk.Label(path_title_row, text="\U0001F4C1", font=("Segoe UI", 9),
                     fg=C_PRIMARY, bg=C_CARD_BG).pack(side="left", padx=(0, 5))
            tk.Label(path_title_row, text="Folder & Aplikasi", font=("Segoe UI", 9, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(side="left")

            tk.Label(wrap, text="Sesuaikan kalau lokasi folder/EBarcode.exe di PC ini berbeda.",
                     font=("Segoe UI", 8), fg=C_MUTED, bg=C_CARD_BG
                     ).pack(fill="x", pady=(0, 6))

            path_fields = [
                ("Folder Data Proses (DIR_PROSES)", DIR_PROSES, "Tempat EXCEL IN/OUT BOX.txt & file status"),
                ("Folder Data Scan (DIR_SCAN)",     DIR_SCAN,   "Tempat file +PRINT IN/OUT LABEL.txt"),
                ("Path EBarcode.exe",               EXE_PATH,   "Lokasi lengkap file EBarcode.exe"),
            ]
            self.path_entries = {}
            for label, value, hint in path_fields:
                block = tk.Frame(wrap, bg=C_CARD_BG)
                block.pack(fill="x", pady=(0, 8))

                tk.Label(block, text=label, font=("Segoe UI", 8, "bold"),
                         fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w")

                holder = tk.Frame(block, bg=C_BORDER)
                holder.pack(fill="x", pady=(3, 0))
                ent = tk.Entry(holder, font=("Segoe UI", 9), relief="flat", bd=0,
                                fg=C_TEXT, bg="#FBFDFF", insertbackground=C_TEXT)
                ent.insert(0, value)
                ent.pack(fill="x", padx=1, pady=1, ipady=4, ipadx=6)
                ent.bind("<FocusIn>", lambda e, h=holder: h.config(bg=C_PRIMARY))
                ent.bind("<FocusOut>", lambda e, h=holder: h.config(bg=C_BORDER))

                tk.Label(block, text=hint, font=("Segoe UI", 7),
                         fg=C_MUTED, bg=C_CARD_BG).pack(anchor="w", pady=(2, 0))

                self.path_entries[label] = ent

            btn_save_path = tk.Button(
                wrap, text="Simpan Folder & Aplikasi", font=("Segoe UI", 8, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=10, pady=5,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._save_paths
            )
            btn_save_path.pack(anchor="e", pady=(0, 2))
            bind_hover(btn_save_path, C_PRIMARY, C_PRIMARY_HOV)

            tk.Frame(wrap, bg=C_BORDER, height=1).pack(fill="x", pady=(8, 8))

            btn_row = tk.Frame(wrap, bg=C_CARD_BG)
            btn_row.pack(fill="x")

            btn_close = tk.Button(
                btn_row, text="Tutup", font=("Segoe UI", 9, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=12, pady=5,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self.destroy
            )
            btn_close.pack(side="right")
            bind_hover(btn_close, C_PRIMARY, C_PRIMARY_HOV)

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

        def _save_paths(self):
            dir_proses = self.path_entries["Folder Data Proses (DIR_PROSES)"].get().strip()
            dir_scan   = self.path_entries["Folder Data Scan (DIR_SCAN)"].get().strip()
            exe_path   = self.path_entries["Path EBarcode.exe"].get().strip()

            try:
                apply_path_config(dir_scan, dir_proses, exe_path)
            except ValueError as e:
                messagebox.showwarning("Folder & Aplikasi", str(e), parent=self)
                return

            ui_log(f"Folder & Aplikasi disimpan: DIR_PROSES={DIR_PROSES}, "
                   f"DIR_SCAN={DIR_SCAN}, EXE_PATH={EXE_PATH}")
            messagebox.showinfo("Folder & Aplikasi", "Pengaturan folder/aplikasi disimpan.", parent=self)

        def _reload_table(self):
            self.tree.delete(*self.tree.get_children())
            for s in STORES:
                status = "Aktif" if s["code"] == ACTIVE_STORE_CODE else ""
                self.tree.insert("", "end", iid=s["code"],
                                  values=(s["code"], s["spreadsheet_id"], s["gid"], status))

        def _selected_code(self):
            sel = self.tree.selection()
            return sel[0] if sel else None

        def _notify_saved(self):
            if self.on_saved:
                self.on_saved()

        def _add_store(self):
            def on_submit(code, sid, gid):
                upsert_store(code, sid, gid, activate=(len(STORES) == 0))
                self._reload_table()
                self._notify_saved()
                ui_log(f"Toko baru ditambahkan: {code}")
            StoreFormDialog(self, "Tambah Toko", on_submit=on_submit)

        def _edit_store(self):
            code = self._selected_code()
            if not code:
                messagebox.showinfo("Edit Toko", "Pilih toko yang mau diedit dulu.", parent=self)
                return
            store = _find_store(code)

            def on_submit(new_code, sid, gid):
                upsert_store(new_code, sid, gid, original_code=code)
                self._reload_table()
                self._notify_saved()
                ui_log(f"Toko '{code}' diperbarui" + (f" menjadi '{new_code}'" if new_code != code else ""))
            StoreFormDialog(self, "Edit Toko", initial=store, on_submit=on_submit)

        def _activate_store(self):
            code = self._selected_code()
            if not code:
                messagebox.showinfo("Jadikan Aktif", "Pilih toko yang mau diaktifkan dulu.", parent=self)
                return
            set_active_store(code)
            self._reload_table()
            self._notify_saved()
            ui_log(f"Toko aktif diubah ke: {STORE_CODE}")

        def _delete_store(self):
            code = self._selected_code()
            if not code:
                messagebox.showinfo("Hapus Toko", "Pilih toko yang mau dihapus dulu.", parent=self)
                return
            if not messagebox.askyesno("Hapus Toko", f"Yakin hapus toko '{code}'?", parent=self):
                return
            try:
                delete_store(code)
            except ValueError as e:
                messagebox.showwarning("Hapus Toko", str(e), parent=self)
                return
            self._reload_table()
            self._notify_saved()
            ui_log(f"Toko '{code}' dihapus")


    class BrownBoxCustomDialog(tk.Toplevel):
        """Dialog input manual untuk cetak Brown Box custom.
        User bisa mengisi BEBERAPA kode box sekaligus (satu kode per baris,
        atau dipisah koma). Setiap kode akan dicetak sebagai satu label
        Brown Box terpisah lewat print_brown_box_labels()."""

        def __init__(self, parent, on_submit):
            super().__init__(parent)
            self.on_submit = on_submit
            self.title("Print Brown Box (Custom)")
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            wrap = tk.Frame(self, bg=C_CARD_BG, padx=22, pady=18)
            wrap.pack(fill="both", expand=True)

            tk.Label(wrap, text="Kode Box", font=("Segoe UI", 11, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w")
            tk.Label(
                wrap,
                text="Masukkan kode box, satu kode per baris\n"
                     "(boleh juga dipisah koma). Bisa banyak sekaligus.",
                font=("Segoe UI", 8), fg=C_MUTED, bg=C_CARD_BG, justify="left"
            ).pack(anchor="w", pady=(2, 10))

            self.txt = tk.Text(
                wrap, width=40, height=10, font=("Consolas", 10),
                bg="#FFFFFF", fg=C_TEXT, relief="flat",
                borderwidth=1, highlightthickness=1, highlightbackground=C_BORDER
            )
            self.txt.pack(fill="both", expand=True)
            self.txt.focus_set()

            self.count_var = tk.StringVar(value="0 kode")
            tk.Label(wrap, textvariable=self.count_var, font=("Segoe UI", 8),
                     fg=C_MUTED, bg=C_CARD_BG).pack(anchor="e", pady=(4, 0))
            self.txt.bind("<KeyRelease>", self._update_count)

            btn_row = tk.Frame(wrap, bg=C_CARD_BG)
            btn_row.pack(fill="x", pady=(14, 0))

            btn_cancel = tk.Button(
                btn_row, text="Batal", font=("Segoe UI", 9, "bold"),
                bg=C_SECONDARY, fg=C_TEXT, relief="flat", padx=14, pady=7,
                activebackground=C_SECONDARY_HOV, activeforeground=C_TEXT,
                cursor="hand2", bd=0, command=self.destroy
            )
            btn_cancel.pack(side="left", expand=True, fill="x", padx=(0, 6))

            self.btn_print = tk.Button(
                btn_row, text="Print", font=("Segoe UI", 9, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=14, pady=7,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._submit
            )
            self.btn_print.pack(side="left", expand=True, fill="x", padx=(6, 0))

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

        def _parse_codes(self) -> list[str]:
            raw = self.txt.get("1.0", "end")
            # Terima pemisah baris baru MAUPUN koma, biar fleksibel cara input user.
            parts = [p.strip() for line in raw.splitlines() for p in line.split(",")]
            # Buang kosong, tapi biarkan duplikat (user mungkin memang mau cetak
            # ulang kode yang sama beberapa kali).
            return [p for p in parts if p]

        def _update_count(self, event=None):
            n = len(self._parse_codes())
            self.count_var.set(f"{n} kode")

        def _submit(self):
            codes = self._parse_codes()
            if not codes:
                messagebox.showwarning(
                    "Print Brown Box", "Masukkan minimal satu kode box.", parent=self
                )
                return
            if len(codes) > 200:
                if not messagebox.askyesno(
                    "Print Brown Box",
                    f"Ada {len(codes)} kode. Ini cukup banyak dan akan memakan waktu lama, "
                    "lanjutkan?", parent=self
                ):
                    return
            self.destroy()
            self.on_submit(codes)


    class GondolaCustomDialog(tk.Toplevel):
        """Dialog input manual untuk cetak Gondola custom.
        User bisa mengisi BEBERAPA kode sekaligus (satu kode per baris,
        atau dipisah koma). Setiap kode akan dicetak sebagai satu label
        Gondola terpisah lewat print_gondola_labels()."""

        def __init__(self, parent, on_submit):
            super().__init__(parent)
            self.on_submit = on_submit
            self.title("Print Gondola (Custom)")
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            wrap = tk.Frame(self, bg=C_CARD_BG, padx=22, pady=18)
            wrap.pack(fill="both", expand=True)

            tk.Label(wrap, text="Kode Gondola", font=("Segoe UI", 11, "bold"),
                     fg=C_TEXT, bg=C_CARD_BG).pack(anchor="w")
            tk.Label(
                wrap,
                text="Masukkan kode, satu kode per baris\n"
                     "(boleh juga dipisah koma). Bisa banyak sekaligus.",
                font=("Segoe UI", 8), fg=C_MUTED, bg=C_CARD_BG, justify="left"
            ).pack(anchor="w", pady=(2, 10))

            self.txt = tk.Text(
                wrap, width=40, height=10, font=("Consolas", 10),
                bg="#FFFFFF", fg=C_TEXT, relief="flat",
                borderwidth=1, highlightthickness=1, highlightbackground=C_BORDER
            )
            self.txt.pack(fill="both", expand=True)
            self.txt.focus_set()

            self.count_var = tk.StringVar(value="0 kode")
            tk.Label(wrap, textvariable=self.count_var, font=("Segoe UI", 8),
                     fg=C_MUTED, bg=C_CARD_BG).pack(anchor="e", pady=(4, 0))
            self.txt.bind("<KeyRelease>", self._update_count)

            btn_row = tk.Frame(wrap, bg=C_CARD_BG)
            btn_row.pack(fill="x", pady=(14, 0))

            btn_cancel = tk.Button(
                btn_row, text="Batal", font=("Segoe UI", 9, "bold"),
                bg=C_SECONDARY, fg=C_TEXT, relief="flat", padx=14, pady=7,
                activebackground=C_SECONDARY_HOV, activeforeground=C_TEXT,
                cursor="hand2", bd=0, command=self.destroy
            )
            btn_cancel.pack(side="left", expand=True, fill="x", padx=(0, 6))

            self.btn_print = tk.Button(
                btn_row, text="Print", font=("Segoe UI", 9, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=14, pady=7,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._submit
            )
            self.btn_print.pack(side="left", expand=True, fill="x", padx=(6, 0))

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

        def _parse_codes(self) -> list[str]:
            raw = self.txt.get("1.0", "end")
            parts = [p.strip() for line in raw.splitlines() for p in line.split(",")]
            return [p for p in parts if p]

        def _update_count(self, event=None):
            n = len(self._parse_codes())
            self.count_var.set(f"{n} kode")

        def _submit(self):
            codes = self._parse_codes()
            if not codes:
                messagebox.showwarning(
                    "Print Gondola", "Masukkan minimal satu kode.", parent=self
                )
                return
            if len(codes) > 200:
                if not messagebox.askyesno(
                    "Print Gondola",
                    f"Ada {len(codes)} kode. Ini cukup banyak dan akan memakan waktu lama, "
                    "lanjutkan?", parent=self
                ):
                    return
            self.destroy()
            self.on_submit(codes)


    class PasscodeDialog(tk.Toplevel):
        """Dialog input kode akses Pengaturan, dilengkapi tombol untuk minta
        kode lewat WhatsApp kalau lupa/tidak punya kodenya."""

        def __init__(self, parent):
            super().__init__(parent)
            self.result = None
            self.title("Pengaturan")
            self.configure(bg=C_CARD_BG)
            self.resizable(False, False)
            self.transient(parent)
            self.grab_set()

            tk.Frame(self, bg=C_HEADER_BG, height=5).pack(fill="x")

            wrap = tk.Frame(self, bg=C_CARD_BG, padx=22, pady=18)
            wrap.pack(fill="both", expand=True)

            tk.Label(wrap, text="Masukkan kode untuk membuka Pengaturan",
                     font=("Segoe UI", 10, "bold"), fg=C_TEXT, bg=C_CARD_BG,
                     wraplength=260, justify="left").pack(anchor="w")

            self.var_code = tk.StringVar(value="")
            entry = tk.Entry(
                wrap, textvariable=self.var_code, show="*", font=("Segoe UI", 11),
                relief="flat", highlightthickness=1, highlightbackground=C_BORDER,
                highlightcolor=C_PRIMARY
            )
            entry.pack(fill="x", pady=(10, 4), ipady=5)
            entry.focus_set()
            entry.bind("<Return>", lambda e: self._submit())

            btn_wa = tk.Button(
                wrap, text="\U0001F4AC Lupa kode? Minta lewat WhatsApp",
                font=("Segoe UI", 8, "underline"), fg=C_PRIMARY, bg=C_CARD_BG,
                relief="flat", bd=0, cursor="hand2",
                activebackground=C_CARD_BG, activeforeground=C_PRIMARY_HOV,
                command=self._open_wa
            )
            btn_wa.pack(anchor="w", pady=(4, 12))

            btn_row = tk.Frame(wrap, bg=C_CARD_BG)
            btn_row.pack(fill="x")

            btn_cancel = tk.Button(
                btn_row, text="Batal", font=("Segoe UI", 9, "bold"),
                bg=C_SECONDARY, fg=C_TEXT, relief="flat", padx=14, pady=7,
                activebackground=C_SECONDARY_HOV, activeforeground=C_TEXT,
                cursor="hand2", bd=0, command=self._cancel
            )
            btn_cancel.pack(side="left", expand=True, fill="x", padx=(0, 6))

            btn_ok = tk.Button(
                btn_row, text="Buka", font=("Segoe UI", 9, "bold"),
                bg=C_PRIMARY, fg="#FFFFFF", relief="flat", padx=14, pady=7,
                activebackground=C_PRIMARY_HOV, activeforeground="#FFFFFF",
                cursor="hand2", bd=0, command=self._submit
            )
            btn_ok.pack(side="left", expand=True, fill="x", padx=(6, 0))

            self.protocol("WM_DELETE_WINDOW", self._cancel)

            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

        def _open_wa(self):
            pesan = urllib.parse.quote(
                "Halo, saya mau minta kode akses menu Pengaturan aplikasi EXIO."
            )
            url = f"https://wa.me/{SETTINGS_WA_NUMBER}?text={pesan}"
            try:
                webbrowser.open(url)
                ui_log("Membuka WhatsApp untuk minta kode akses Pengaturan.")
            except Exception as e:
                log.warning(f"Gagal membuka link WhatsApp: {e}")
                messagebox.showinfo(
                    "WhatsApp",
                    f"Silakan hubungi WA berikut secara manual:\n{SETTINGS_WA_NUMBER}",
                    parent=self
                )

        def _submit(self):
            self.result = self.var_code.get()
            self.destroy()

        def _cancel(self):
            self.result = None
            self.destroy()


    class App:
        def __init__(self, root):
            self.root = root
            self.root.title(f"KS UPDATE - {STORE_CODE} (EXINOUT + Print)")
            self.root.configure(bg=C_BG)
            self.root.resizable(False, False)
            # Hilangkan title bar bawaan OS (teks judul + tombol exit/minimize bawaan)
            self.root.overrideredirect(True)
            self.root.configure(highlightthickness=1, highlightbackground=C_BORDER,
                                 highlightcolor=C_BORDER)

            header_bar = tk.Frame(root, bg=C_HEADER_BG, height=10, cursor="fleur")
            header_bar.pack(fill="x")
            # Drag window lewat strip atas, karena title bar OS sudah dihilangkan
            header_bar.bind("<Button-1>", self._drag_start)
            header_bar.bind("<B1-Motion>", self._drag_move)

            header_inner = tk.Frame(header_bar, bg=C_HEADER_BG, padx=12, pady=6)
            header_inner.pack(fill="x")
            header_inner.bind("<Button-1>", self._drag_start)
            header_inner.bind("<B1-Motion>", self._drag_move)

            self.header_lbl = tk.Label(
                header_inner, text=f"EXIO v2 - {STORE_CODE}",
                font=("Segoe UI", 10, "bold"), fg=C_HEADER_FG, bg=C_HEADER_BG
            )
            self.header_lbl.pack(side="left")
            self.header_lbl.bind("<Button-1>", self._drag_start)
            self.header_lbl.bind("<B1-Motion>", self._drag_move)

            self.btn_settings = tk.Button(
                header_inner, text="\u2699", font=("Segoe UI", 10),
                bg=C_HEADER_BG, fg=C_HEADER_FG, relief="flat", padx=4, pady=0,
                activebackground=C_GEAR_HOV, activeforeground=C_HEADER_FG,
                cursor="hand2", bd=0, highlightthickness=0,
                command=self.open_settings
            )
            self.btn_settings.pack(side="right")

            btn_frame = tk.Frame(root, bg=C_BG, padx=10, pady=10)
            btn_frame.pack(fill="x")

            BTN_OPTS = dict(font=("Segoe UI", 9, "bold"), fg=C_TEXT,
                             relief="flat", width=13, height=2, bd=0, cursor="hand2")

            self.btn_update = tk.Button(
                btn_frame, text="Sync IN/OUT",
                bg=C_PRIMARY, activebackground=C_PRIMARY_HOV, activeforeground=C_TEXT,
                command=lambda: self.on_click("AUTO"), **BTN_OPTS
            )
            self.btn_update.pack(side="left", padx=(0, 6), expand=True, fill="x")
            bind_hover(self.btn_update, C_PRIMARY, C_PRIMARY_HOV)

            self.btn_recon = tk.Button(
                btn_frame, text="Sync Recon",
                bg=C_RECON, activebackground=C_RECON_HOV, activeforeground=C_TEXT,
                command=lambda: self.on_click("RECON"), **BTN_OPTS
            )
            self.btn_recon.pack(side="left", padx=6, expand=True, fill="x")
            bind_hover(self.btn_recon, C_RECON, C_RECON_HOV)

            self.btn_repeat = tk.Button(
                btn_frame, text="Reprint",
                bg=C_SECONDARY, activebackground=C_SECONDARY_HOV, activeforeground=C_TEXT,
                command=lambda: self.on_click("REPEAT"), **BTN_OPTS
            )
            self.btn_repeat.pack(side="left", padx=(6, 0), expand=True, fill="x")
            bind_hover(self.btn_repeat, C_SECONDARY, C_SECONDARY_HOV)

            btn_frame2 = tk.Frame(root, bg=C_BG, padx=10)
            btn_frame2.pack(fill="x", pady=(0, 6))

            self.btn_gondola = tk.Button(
                btn_frame2, text="Print Gondola (Custom)",
                font=("Segoe UI", 9, "bold"), fg=C_TEXT,
                relief="flat", height=2, bd=0, cursor="hand2",
                bg=C_RECON, activebackground=C_RECON_HOV, activeforeground=C_TEXT,
                command=self.open_gondola_dialog
            )
            self.btn_gondola.pack(side="left", padx=(0, 6), expand=True, fill="x")
            bind_hover(self.btn_gondola, C_RECON, C_RECON_HOV)

            self.btn_brown_box = tk.Button(
                btn_frame2, text="Print Brown Box (Custom)",
                font=("Segoe UI", 9, "bold"), fg=C_TEXT,
                relief="flat", height=2, bd=0, cursor="hand2",
                bg=C_RECON, activebackground=C_RECON_HOV, activeforeground=C_TEXT,
                command=self.open_brown_box_dialog
            )
            self.btn_brown_box.pack(side="left", expand=True, fill="x")
            bind_hover(self.btn_brown_box, C_RECON, C_RECON_HOV)

            btn_frame3 = tk.Frame(root, bg=C_BG, padx=10)
            btn_frame3.pack(fill="x", pady=(0, 6))

            self.PAUSE_TEXT  = "Jeda\n(Ctrl+Alt+P)"
            self.RESUME_TEXT = "Lanjut\n(Ctrl+Alt+P)"
            self._status_before_pause = ""

            self.btn_pause = tk.Button(
                btn_frame3, text=self.PAUSE_TEXT, font=("Segoe UI", 8, "bold"), fg=C_TEXT,
                relief="flat", height=2, bd=0, cursor="hand2", state="disabled",
                bg=C_PAUSE, activebackground=C_PAUSE_HOV, activeforeground=C_TEXT,
                command=self.toggle_pause
            )
            self.btn_pause.pack(side="left", expand=True, fill="x", padx=(0, 6))

            self.btn_stop = tk.Button(
                btn_frame3, text="STOP\n(Ctrl+Alt+S)", font=("Segoe UI", 8, "bold"), fg=C_TEXT,
                relief="flat", height=2, bd=0, cursor="hand2", state="disabled",
                bg=C_STOP, activebackground=C_STOP_HOV, activeforeground=C_TEXT,
                command=self.stop_run
            )
            self.btn_stop.pack(side="left", expand=True, fill="x", padx=(6, 0))

            self.status_var = tk.StringVar(value="Siap.")
            self.status_lbl = tk.Label(
                root, textvariable=self.status_var, font=("Segoe UI", 8, "bold"),
                fg=C_STATUS, bg=C_BG
            )
            self.status_lbl.pack(pady=(0, 4))

            self.log_widget = tk.Text(
                root, width=50, height=1, state="disabled", wrap="none",
                bg=C_LOG_BG, fg=C_LOG_FG, font=("Consolas", 8), relief="flat",
                borderwidth=1, highlightthickness=1, highlightbackground=C_BORDER,
                insertbackground=C_LOG_FG
            )
            self.log_widget.pack(fill="x", padx=10, pady=(0, 6))

            footer_frame = tk.Frame(root, bg=C_BG, padx=10)
            footer_frame.pack(fill="x", pady=(0, 10))

            self.btn_exit = tk.Button(
                footer_frame, text="Exit", font=("Segoe UI", 8, "bold"),
                bg=C_SECONDARY, fg=C_TEXT, relief="flat", padx=14, pady=4,
                activebackground=C_SECONDARY_HOV, activeforeground=C_TEXT,
                cursor="hand2", bd=0, command=self.root.destroy
            )
            self.btn_exit.pack(side="right")
            bind_hover(self.btn_exit, C_SECONDARY, C_SECONDARY_HOV)

            handler = TextHandler(self.log_widget)
            handler.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S"))
            log.addHandler(handler)

            ui_log("Aplikasi siap.")

            if start_global_hotkeys(
                on_pause=lambda: self.root.after(0, self.toggle_pause),
                on_stop=lambda: self.root.after(0, self.stop_run),
            ):
                log.info("Hotkey aktif: Ctrl+Alt+P = Jeda/Lanjut, Ctrl+Alt+S = Stop")

            self._center_window()

            # Auto-close otomatis 5 menit setelah SEMUA proses selesai
            # (atau 5 menit sejak aplikasi dibuka kalau tidak ada proses sama sekali)
            self.AUTO_CLOSE_MS = 5 * 60 * 1000  # 5 menit
            self._auto_close_job = None
            self._schedule_auto_close()

        def _auto_close(self):
            ui_log("Aplikasi ditutup otomatis (5 menit setelah proses selesai, tidak ada aktivitas lain).")
            try:
                self.root.destroy()
            except Exception:
                pass

        def _cancel_auto_close(self):
            if self._auto_close_job is not None:
                try:
                    self.root.after_cancel(self._auto_close_job)
                except Exception:
                    pass
                self._auto_close_job = None

        def _schedule_auto_close(self):
            self._cancel_auto_close()
            self._auto_close_job = self.root.after(self.AUTO_CLOSE_MS, self._auto_close)

        def _center_window(self):
            self.root.update_idletasks()
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = (sw - w) // 2
            y = (sh - h) // 2
            self.root.geometry(f"{w}x{h}+{x}+{y}")

        def _drag_start(self, event):
            self._drag_x = event.x
            self._drag_y = event.y

        def _drag_move(self, event):
            x = self.root.winfo_x() + (event.x - self._drag_x)
            y = self.root.winfo_y() + (event.y - self._drag_y)
            self.root.geometry(f"+{x}+{y}")

        def open_settings(self):
            dlg = PasscodeDialog(self.root)
            self.root.wait_window(dlg)
            code = dlg.result
            if code is None:
                return  # dibatalkan
            if code.strip() != SETTINGS_PASSCODE:
                messagebox.showerror(
                    "Pengaturan",
                    "Kode salah.\n\nKalau lupa kode, klik ikon \u2699 lagi lalu "
                    "pilih 'Lupa kode? Minta lewat WhatsApp'.",
                    parent=self.root
                )
                ui_log("Percobaan buka Pengaturan gagal (kode salah).")
                return
            SettingsDialog(self.root, on_saved=self._refresh_title)

        def _refresh_title(self):
            self.root.title(f"KS UPDATE - {STORE_CODE} (EXINOUT + Print)")
            self.header_lbl.config(text=f"KS - {STORE_CODE}")

        def open_gondola_dialog(self):
            GondolaCustomDialog(self.root, on_submit=self._run_gondola_custom)

        def _set_running_ui(self, running: bool):
            state = "normal" if running else "disabled"
            self.btn_pause.config(state=state, text=self.PAUSE_TEXT)
            self.btn_stop.config(state=state)

        def toggle_pause(self):
            """Jeda / lanjutkan otomasi EBarcode (aman kapan saja selama proses berjalan)."""
            if not control.active or control.stop_requested:
                return
            if control.paused:
                control.set_paused(False)
                self.btn_pause.config(text=self.PAUSE_TEXT)
                self.status_var.set(self._status_before_pause or "Memproses...")
                ui_log("Proses DILANJUTKAN.")
            else:
                control.set_paused(True)
                self._status_before_pause = self.status_var.get()
                self.btn_pause.config(text=self.RESUME_TEXT)
                self.status_var.set("DIJEDA - tekan Lanjut (Ctrl+Alt+P) untuk meneruskan")
                ui_log("Proses DIJEDA. Tekan Lanjut untuk meneruskan atau STOP untuk menghentikan.")

        def stop_run(self):
            """Hentikan otomasi. Data yang belum selesai TIDAK hilang: muncul lagi di Sync berikutnya."""
            if not control.active or control.stop_requested:
                return
            control.request_stop()
            self.btn_pause.config(state="disabled")
            self.btn_stop.config(state="disabled")
            self.status_var.set("Menghentikan proses...")
            ui_log("STOP diterima, menghentikan otomasi EBarcode...")

        def open_brown_box_dialog(self):
            BrownBoxCustomDialog(self.root, on_submit=self._run_brown_box_custom)

        def _run_gondola_custom(self, box_codes: list[str]):
            self._cancel_auto_close()
            control.begin()
            self._set_running_ui(True)
            self.btn_update.config(state="disabled")
            self.btn_recon.config(state="disabled")
            self.btn_repeat.config(state="disabled")
            self.btn_gondola.config(state="disabled")
            self.btn_brown_box.config(state="disabled")
            self.status_var.set(f"Mencetak {len(box_codes)} Gondola...")

            self.log_widget.configure(state="normal")
            self.log_widget.delete("1.0", "end")
            self.log_widget.configure(state="disabled")

            thread = threading.Thread(
                target=self._run_gondola_thread, args=(box_codes,), daemon=True
            )
            thread.start()

        def _run_gondola_thread(self, box_codes: list[str]):
            try:
                ui_log(f"GONDOLA CUSTOM - {len(box_codes)} kode akan dicetak: {box_codes}")
                success_codes, failed_codes = print_gondola_labels(box_codes)
                ok = len(failed_codes) == 0
                message = (
                    f"Print Gondola (Custom) selesai.\n\n"
                    f"\u25CF Berhasil ({len(success_codes)}): {', '.join(success_codes) or '-'}\n"
                    f"\u25CF Gagal ({len(failed_codes)})   : {', '.join(failed_codes) or '-'}"
                )
            except StopRequested:
                ok, message = False, "Print Gondola DIHENTIKAN manual oleh pengguna."
            except Exception as e:
                ok, message = False, f"Error tidak terduga saat print Gondola: {e}"
                log.error(message)
            stopped = (not ok) and control.stop_requested
            control.end()

            def finish():
                self.status_var.set("Siap." if ok else ("Dihentikan." if stopped else "Terjadi kesalahan."))
                self.btn_update.config(state="normal")
                self.btn_recon.config(state="normal")
                self.btn_repeat.config(state="normal")
                self.btn_gondola.config(state="normal")
                self.btn_brown_box.config(state="normal")
                self._set_running_ui(False)
                show_notice(success=ok, message=message, stopped=stopped)
                self._schedule_auto_close()

            self.root.after(0, finish)

        def _run_brown_box_custom(self, box_codes: list[str]):
            self._cancel_auto_close()
            control.begin()
            self._set_running_ui(True)
            self.btn_update.config(state="disabled")
            self.btn_recon.config(state="disabled")
            self.btn_repeat.config(state="disabled")
            self.btn_gondola.config(state="disabled")
            self.btn_brown_box.config(state="disabled")
            self.status_var.set(f"Mencetak {len(box_codes)} Brown Box...")

            self.log_widget.configure(state="normal")
            self.log_widget.delete("1.0", "end")
            self.log_widget.configure(state="disabled")

            thread = threading.Thread(
                target=self._run_brown_box_thread, args=(box_codes,), daemon=True
            )
            thread.start()

        def _run_brown_box_thread(self, box_codes: list[str]):
            try:
                ui_log(f"BROWN BOX CUSTOM - {len(box_codes)} kode akan dicetak: {box_codes}")
                success_codes, failed_codes = print_brown_box_labels(box_codes)
                ok = len(failed_codes) == 0
                message = (
                    f"Print Brown Box (Custom) selesai.\n\n"
                    f"\u25CF Berhasil ({len(success_codes)}): {', '.join(success_codes) or '-'}\n"
                    f"\u25CF Gagal ({len(failed_codes)})   : {', '.join(failed_codes) or '-'}"
                )
            except StopRequested:
                ok, message = False, "Print Brown Box DIHENTIKAN manual oleh pengguna."
            except Exception as e:
                ok, message = False, f"Error tidak terduga saat print Brown Box: {e}"
                log.error(message)
            stopped = (not ok) and control.stop_requested
            control.end()

            def finish():
                self.status_var.set("Siap." if ok else ("Dihentikan." if stopped else "Terjadi kesalahan."))
                self.btn_update.config(state="normal")
                self.btn_recon.config(state="normal")
                self.btn_repeat.config(state="normal")
                self.btn_gondola.config(state="normal")
                self.btn_brown_box.config(state="normal")
                self._set_running_ui(False)
                show_notice(success=ok, message=message, stopped=stopped)
                self._schedule_auto_close()

            self.root.after(0, finish)

        def on_click(self, mode: str):
            self._cancel_auto_close()
            control.begin()
            self._set_running_ui(True)
            self.btn_update.config(state="disabled")
            self.btn_recon.config(state="disabled")
            self.btn_repeat.config(state="disabled")
            self.btn_gondola.config(state="disabled")
            self.btn_brown_box.config(state="disabled")
            if mode == "REPEAT":
                self.status_var.set("Mengulang proses EBarcode (tanpa ambil data GSheet)...")
            elif mode == "RECON":
                self.status_var.set("Memproses KS RECON (Keep Stock vs Master)...")
            else:
                self.status_var.set("Memproses UPDATE KS...")

            self.log_widget.configure(state="normal")
            self.log_widget.delete("1.0", "end")
            self.log_widget.configure(state="disabled")

            thread = threading.Thread(target=self._run_thread, args=(mode,), daemon=True)
            thread.start()

        def _run_thread(self, mode: str):
            try:
                if mode == "REPEAT":
                    ok, message = run_flow_repeat()
                elif mode == "RECON":
                    ok, message = run_flow_recon()
                else:
                    ok, message = run_flow_auto()
            except StopRequested:
                ok, message = False, ("Proses DIHENTIKAN manual oleh pengguna.\n\n"
                                      "Data yang belum selesai tidak hilang: akan muncul lagi "
                                      "saat Sync berikutnya (atau tekan Reprint).")
            except Exception as e:
                ok, message = False, f"Error tidak terduga: {e}"
                log.error(message)
            stopped = (not ok) and control.stop_requested
            control.end()

            def finish():
                self.status_var.set("Siap." if ok else ("Dihentikan." if stopped else "Terjadi kesalahan."))
                self.btn_update.config(state="normal")
                self.btn_recon.config(state="normal")
                self.btn_repeat.config(state="normal")
                self.btn_gondola.config(state="normal")
                self.btn_brown_box.config(state="normal")
                self._set_running_ui(False)
                show_notice(success=ok, message=message, stopped=stopped)
                self._schedule_auto_close()

            self.root.after(0, finish)


    # ============================================================
    # MAIN
    # ============================================================

def main():
    log.info("=" * 60)
    log.info(f"  KS UPDATE - {STORE_CODE} (EXINOUT + PrintAuto Terintegrasi)")
    log.info(f"  Sumber data (proses)   : {DIR_PROSES}")
    log.info(f"  Sumber data (scan/print): {DIR_SCAN}")
    log.info("  Tidak ada jam operasional / penantian file otomatis.")
    log.info("  Semua klik otomasi EBarcode berbasis PostMessage (safe click).")
    if PYVDA_OK:
        log.info("  Virtual Desktop: AKTIF (pyvda terinstall)")
    else:
        log.info("  Virtual Desktop: TIDAK AKTIF — jalankan: pip install pyvda")
    log.info("=" * 60)

    if not HAS_TK:
        log.error("tkinter tidak tersedia, tidak bisa menampilkan GUI.")
        sys.exit(1)

    root = tk.Tk()

    if not STORE_SELECTED:
        # Root sengaja TIDAK di-withdraw() di sini. Kalau root masih hidden saat
        # dialog (transient ke root) dibuka, di sebagian PC (terutama Python dari
        # Microsoft Store / lingkungan sandbox) grab_set()/wait_visibility() bisa
        # gagal atau bahkan hang tanpa error, sehingga aplikasi terlihat 'macet'.
        # Root dibiarkan tampil sebentar (kosong) di belakang dialog; setelah App
        # dibuat, root langsung terisi tampilan utama seperti biasa.
        try:
            dlg = StoreSelectDialog(root)
            root.wait_window(dlg)
            log.info(f"  Toko aktif dipilih: {STORE_CODE}")
        except Exception as e:
            log.error(f"Dialog pilih toko gagal dibuka: {e}")
            log.error("Melanjutkan tanpa toko aktif terkonfirmasi — "
                       "silakan pilih toko lewat menu Pengaturan.")
    else:
        root.withdraw()

    root.deiconify()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.exception("Aplikasi berhenti karena error tak tertangani:")
        import traceback
        _write_crash_fallback("Aplikasi berhenti karena error tak tertangani:\n" + traceback.format_exc())
        raise