"""
UNIFIED AUTO — Satu skrip gabungan
===================================
Urutan proses (otomatis, jeda 5 detik antar-tahap):
  TAHAP 1: XShelf         -> logikanya PERSIS SAMA seperti XShelf.py asli
                              (QubeDrive upload_data -> download_data, RUN#1 & RUN#2 export SHELF)
  TAHAP 2: CheckDisc       -> headless, baca data, bandingkan dgn snapshot,
                              export PRICE CHANGE / CHANGE DESCRIPTION / NO SHELF
                              ke D:\\Scanning (tanpa GUI, tanpa notifikasi)
                              -> kalau TIDAK ADA perubahan, skrip langsung TUTUP
                                 (TAHAP 3 tidak dijalankan)
  TAHAP 3: PrintAuto24     -> Step 1 (buka EBarcode), Step 2 (klik Print),
                              Step 3 (klik Zebex), Step 4 (TdlgData: un-tick
                              semua lalu centang HANYA file bertanggal HARI INI
                              dari hasil TAHAP 2) -> Extract -> Print -> tutup

Setelah semua tahap sukses, aplikasi otomatis exit setelah 10 detik.

Requires: pip install pywinauto
Jalankan : python unified_auto.py
"""

import os
import sys
import time
import json
import re
import glob
import csv
import logging
import ctypes
import subprocess
from datetime import datetime, timedelta
from collections import defaultdict

from pywinauto import Application, Desktop

# ============================================================
# KONFIGURASI GLOBAL
# ============================================================
EXE_PATH        = r"D:\QasDev\QubeV10\BackEnd\EBarcode.exe"
QUBEDRIVE_EXE   = r"D:\QasDev\QubeV10\BackEnd\QubeDrive.exe"
QUBEDRIVE_UPLOAD_ARGS = ["-Jupload_data"]
QUBEDRIVE_UPLOAD_WAIT = 2 * 60     # batas MAKSIMAL aman (menit->detik) menunggu
                                     # QubeDrive.exe (upload_data) tertutup sendiri.
                                     # Begitu proses-nya exit, langsung lanjut
                                     # tanpa perlu menunggu sampai batas ini.
QUBEDRIVE_ARGS  = ["-Jdownload_data"]
QUBEDRIVE_WAIT  = 5 * 60     # batas MAKSIMAL aman menunggu QubeDrive.exe
                               # (download_data) tertutup sendiri. Begitu proses
                               # exit, langsung lanjut tanpa nunggu batas ini.
TIMEOUT         = 15         # detik, timeout tunggu window
WAIT_TIMER      = 20 * 60    # 20 menit, tunggu setelah klik Search di XShelf

# --- Cek freshness data hasil download QubeDrive ---
# Dicek dari TANGGAL MODIFIKASI file (bukan dari nama file), supaya tidak
# tergantung pola penamaan tertentu.
QUBEDRIVE_UPDATED_FOLDER          = r"D:\QasDev\QubeV10\BackEnd\Data\QubeDrive\In\IDHQ\Updated"
QUBEDRIVE_FRESHNESS_WAIT_MAX      = 5 * 60   # tunggu maksimal 5 menit
QUBEDRIVE_FRESHNESS_POLL_INTERVAL = 20       # cek folder tiap 20 detik
QUBEDRIVE_RETRY_WAIT_SECONDS      = 6 * 60 * 60   # 6 jam - kalau data belum
                                                    # update, tunggu 6 jam lalu
                                                    # retry download (tanpa upload)

CHECKDISC_INPUT_PATH  = r"D:\QasDev\QubeV10\BackEnd\Data\Seuic\02\02.txt"
CHECKDISC_OUTPUT_PATH = r"D:\Scanning"
PRICE_HISTORY_FILE    = "price_history.json"

# --- Filter/urutan SKU per shelf (opsional) ---
# File CSV/TXT berisi baris: Shelf,SKU,NomorUrut
# Contoh:
#   AG01-01,1234566,1
#   AG01-01,1234567,2
#   AG01-02,1231231,1
# Kalau kolom NomorUrut KOSONG (atau baris/SKU tidak ada di file ini sama
# sekali), item itu TIDAK difilter urutan - tetap ikut proses seperti biasa,
# hanya diurutkan berdasarkan shelf saja (taruh paling akhir di shelf-nya).
SKU_ORDER_FILE = r"D:\Scanning\SKU_ORDER.csv"

# --- QTY MINUS (qty system < 0) & label COLLECTING LABEL (menu list nomor 4) ---
QTY_MINUS_BASENAME        = "QTY MINUS"
COLLECTING_LABEL_NAME     = "COLLECTING LABEL"
COLLECTING_LABEL_ITEM_INDEX = 3   # nomor 4 di Select Label (0-based index = nomor - 1).
                                   # Sesuaikan angka ini kalau posisi COLLECTING LABEL
                                   # di daftar Select Label EBarcode berbeda.
DATE_FORMAT            = "%d/%m/%Y"
EXPIRED_LOOKBACK_DAYS  = 1   # rentang "baru selesai promo" -> PRICE ADJUSTMENT (tercetak 1x saja, di hari promo berakhir)
START_LOOKBACK_DAYS    = 1   # rentang "promo baru mulai" -> ikut PRICE ADJUSTMENT (tercetak 1x saja, di hari promo mulai)

GAP_SECONDS         = 5     # jeda antar tahap
EXIT_DELAY_SUCCESS  = 10    # tunggu sebelum exit setelah semua sukses

BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)))
LOG_FILE = os.path.join(BASE_DIR, "unified_auto_log.txt")
PROGRESS_FILE = os.path.join(BASE_DIR, "progress_status.json")


def _reset_log_if_new_day():
    """Log di-reset (dihapus) tiap ganti hari, supaya tidak menumpuk.
    Kalau masih hari yang sama (mis. dijalankan ulang manual setelah
    macet), log HARI ITU tetap dipertahankan/di-append, tidak dihapus."""
    if os.path.exists(LOG_FILE):
        try:
            mtime_date = datetime.fromtimestamp(os.path.getmtime(LOG_FILE)).date()
            if mtime_date != datetime.now().date():
                os.remove(LOG_FILE)
        except Exception:
            pass


_reset_log_if_new_day()

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger()


# ============================================================================
# ================  PROGRESS TRACKING (buat resume manual)  =================
# Supaya kalau skrip di-klik manual kapan saja: kalau semua tahap utk hari
# ini sudah selesai, tidak dijalankan ulang; kalau proses macet di suatu
# tahap, tahap yang sudah sukses dilewati dan langsung lanjut dari tahap
# yang gagal sampai selesai.
# ============================================================================
def load_progress():
    today = datetime.now().strftime("%Y-%m-%d")
    default = {"date": today, "stages": {}}
    if not os.path.exists(PROGRESS_FILE):
        return default
    try:
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("date") != today:
            # Sudah ganti hari -> progress kemarin tidak berlaku lagi
            return default
        if "stages" not in data:
            data["stages"] = {}
        return data
    except Exception as e:
        log.warning(f"[Progress] Gagal baca {PROGRESS_FILE}, mulai dari awal: {e}")
        return default


def save_progress(progress):
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress, f, indent=2, ensure_ascii=False)
    except Exception as e:
        log.warning(f"[Progress] Gagal simpan {PROGRESS_FILE}: {e}")


def mark_stage_done(progress, stage_name):
    progress.setdefault("stages", {})[stage_name] = True
    save_progress(progress)


# ============================================================================
# ============================  TAHAP 1: XSHELF  ============================
# (logika identik dengan XShelf.py asli — hanya diberi prefix xs_ agar tidak
#  bentrok nama fungsi dengan tahap CheckDisc / PrintAuto24 di file yang sama)
# ============================================================================
def xs_wait_win(class_name, timeout=TIMEOUT):
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="win32").windows(class_name=class_name)
        if wins:
            return wins[0]
        time.sleep(0.25)
    return None


def xs_get_app():
    return Application(backend="win32").connect(path=EXE_PATH)


# ============================================================================
# ============  KLIK BERBASIS WINDOW MESSAGE (immun kursor digeser)  ========
# pywinauto .click() bawaan menggerakkan kursor mouse ASLI ke posisi
# tombol lalu klik di titik itu - kalau user/proses lain menggeser kursor
# di tengah-tengah proses ini, klik bisa meleset/gagal.
# click_control_message() sebaliknya mengirim pesan klik LANGSUNG ke
# handle window control (BM_CLICK untuk tombol, fallback WM_LBUTTONDOWN/UP
# lewat PostMessage), tidak menggerakkan kursor sama sekali, jadi tidak
# terganggu walau mouse fisik sedang dipakai/digeser di tempat lain.
# ============================================================================
_WM_LBUTTONDOWN = 0x0201
_WM_LBUTTONUP   = 0x0202
_BM_CLICK       = 0x00F5


def window_already_open(class_name):
    """Cek SEKALI SAJA (instant, TIDAK polling/menunggu) apakah window
    dengan class_name tsb sudah terbuka saat ini. Dipakai supaya di setiap
    step, kalau ternyata EBarcode sudah ada di step berikutnya (misal
    diklik manual sebelumnya, atau sisa proses lama), script LANGSUNG
    lanjut/klik ke step sesudahnya tanpa menunggu window yang sebenarnya
    sudah terbuka."""
    wins = Desktop(backend="win32").windows(class_name=class_name)
    return wins[0] if wins else None


def click_control_message(control, retries=2):
    """Klik sebuah control (pywinauto wrapper) lewat window message,
    tanpa menggerakkan kursor mouse fisik sama sekali.

    FIX DEADLOCK: sebelumnya BM_CLICK dikirim pakai SendMessageW, yaitu
    panggilan SINKRON/BLOCKING - Windows baru mengembalikan kontrol ke
    Python setelah window PENERIMA selesai memproses klik itu. Untuk
    tombol seperti "Print" (TButton3) yang handler OnClick-nya memanggil
    ShowModal (membuka TfrmPrint sbg dialog modal), ShowModal MASUK KE
    MESSAGE-LOOP SENDIRI yang baru keluar setelah dialog itu ditutup.
    Akibatnya SendMessageW hang menunggu TfrmPrint ditutup, padahal
    TfrmPrint baru bisa ditutup oleh LANGKAH SKRIP BERIKUTNYA -> deadlock
    persis seperti yang terlihat: window sudah kebuka tapi skrip masih
    "menunggu" klik selesai.

    Sekarang PostMessageW (non-blocking / fire-and-forget) dipakai
    sebagai cara UTAMA - skrip mengirim pesan lalu LANGSUNG lanjut tanpa
    menunggu window tujuan selesai memproses, sehingga tidak pernah hang
    walau klik itu memicu window/dialog modal baru."""
    last_err = None
    for attempt in range(retries + 1):
        try:
            hwnd = control.handle
        except Exception as e:
            last_err = e
            break

        # Cara UTAMA (non-blocking): POST BM_CLICK, jangan SEND.
        # PostMessage cukup untuk memicu OnClick tombol Delphi
        # (TButton/TcxButton/TBitBtn) tanpa memblokir skrip menunggu
        # window/dialog baru yang mungkin dibuka oleh handler-nya.
        try:
            ok = ctypes.windll.user32.PostMessageW(hwnd, _BM_CLICK, 0, 0)
            if ok:
                return True
            last_err = "PostMessageW(BM_CLICK) return 0"
        except Exception as e:
            last_err = e

        # Fallback: kirim WM_LBUTTONDOWN + WM_LBUTTONUP ke titik tengah
        # control (koordinat client, bukan koordinat layar), lewat
        # PostMessage juga - tetap non-blocking, tidak menggerakkan
        # kursor asli.
        try:
            rect = control.rectangle()
            cx = max((rect.right - rect.left) // 2, 1)
            cy = max((rect.bottom - rect.top) // 2, 1)
            lparam = (cy << 16) | (cx & 0xFFFF)
            ctypes.windll.user32.PostMessageW(hwnd, _WM_LBUTTONDOWN, 1, lparam)
            time.sleep(0.03)
            ctypes.windll.user32.PostMessageW(hwnd, _WM_LBUTTONUP, 0, lparam)
            return True
        except Exception as e:
            last_err = e
            time.sleep(0.15)

    log.warning(f"  [click_control_message] Gagal klik control setelah beberapa percobaan: {last_err}")
    return False


def xs_safe_set_focus(control):
    try:
        control.set_focus()
        return True
    except Exception as e:
        log.warning(f"[XShelf] set_focus dilewati: {e}")
        return False


def xs_kill_ebarcode():
    log.warning("  [XShelf][KILL] Menutup EBarcode...")
    for cls in ["TdlgData", "TdlgExportSHELF", "TdlgStoreSHELF", "TMessageForm", "TfrmPrint", "TfrmBarcode"]:
        wins = Desktop(backend="win32").windows(class_name=cls)
        for w in wins:
            try:
                xs_safe_set_focus(w)
                w.close()
                time.sleep(0.15)
            except Exception as e:
                log.warning(f"  [XShelf][KILL] Gagal tutup {cls}: {e}")

    time.sleep(0.75)
    wins = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if wins:
        log.warning("  [XShelf][KILL] EBarcode masih ada, paksa taskkill...")
        try:
            subprocess.run(["taskkill", "/F", "/IM", "EBarcode.exe"], capture_output=True, timeout=5)
        except Exception as e:
            log.warning(f"  [XShelf][KILL] taskkill error: {e}")
    else:
        log.warning("  [XShelf][KILL] EBarcode sudah tertutup")


def xs_run_qubedrive_upload() -> bool:
    log.info("[XShelf] STEP 0-PRE - Jalankan QubeDrive.exe -Jupload_data")
    try:
        proc = subprocess.Popen([QUBEDRIVE_EXE] + QUBEDRIVE_UPLOAD_ARGS)
        log.info(f"  OK: QubeDrive dijalankan ({QUBEDRIVE_EXE} {' '.join(QUBEDRIVE_UPLOAD_ARGS)})")
    except Exception as e:
        log.error(f"*** FAIL: Gagal menjalankan QubeDrive.exe -Jupload_data: {e}")
        return False

    log.info(f"[XShelf] STEP 0-PRE b - Menunggu QubeDrive.exe (upload_data) tertutup sendiri "
              f"(proses selesai = exit), maksimal {QUBEDRIVE_UPLOAD_WAIT // 60} menit...")
    start_wait = time.time()
    last_logged_min = -1
    while True:
        if proc.poll() is not None:
            log.info(f"  OK: QubeDrive (upload_data) sudah tertutup/exit setelah "
                      f"~{int(time.time() - start_wait)} detik, lanjut proses berikutnya.")
            return True
        elapsed = time.time() - start_wait
        if elapsed >= QUBEDRIVE_UPLOAD_WAIT:
            log.warning(f"  [PERINGATAN] QubeDrive (upload_data) belum tertutup setelah "
                        f"{QUBEDRIVE_UPLOAD_WAIT // 60} menit (batas aman). Lanjut proses "
                        f"berikutnya saja supaya tidak macet total.")
            return True
        elapsed_min = int(elapsed // 60)
        if elapsed_min != last_logged_min:
            log.info(f"  ... QubeDrive (upload_data) masih berjalan, menunggu tertutup sendiri "
                      f"(sudah {elapsed_min} menit)")
            last_logged_min = elapsed_min
        time.sleep(2)


def xs_run_qubedrive_download() -> bool:
    log.info("[XShelf] STEP 0 - Jalankan QubeDrive.exe -Jdownload_data")
    try:
        proc = subprocess.Popen([QUBEDRIVE_EXE] + QUBEDRIVE_ARGS)
        log.info(f"  OK: QubeDrive dijalankan ({QUBEDRIVE_EXE} {' '.join(QUBEDRIVE_ARGS)})")
    except Exception as e:
        log.error(f"*** FAIL: Gagal menjalankan QubeDrive.exe: {e}")
        return False

    log.info(f"[XShelf] STEP 0b - Menunggu QubeDrive.exe (download_data) tertutup sendiri "
              f"(proses selesai = exit), maksimal {QUBEDRIVE_WAIT // 60} menit...")
    start_wait = time.time()
    last_logged_min = -1
    while True:
        if proc.poll() is not None:
            log.info(f"  OK: QubeDrive (download_data) sudah tertutup/exit setelah "
                      f"~{int(time.time() - start_wait)} detik, lanjut proses berikutnya.")
            return True
        elapsed = time.time() - start_wait
        if elapsed >= QUBEDRIVE_WAIT:
            log.warning(f"  [PERINGATAN] QubeDrive (download_data) belum tertutup setelah "
                        f"{QUBEDRIVE_WAIT // 60} menit (batas aman). Lanjut proses "
                        f"berikutnya saja supaya tidak macet total.")
            return True
        elapsed_min = int(elapsed // 60)
        if elapsed_min != last_logged_min:
            log.info(f"  ... QubeDrive (download_data) masih berjalan, menunggu tertutup sendiri "
                      f"(sudah {elapsed_min} menit)")
            last_logged_min = elapsed_min
        time.sleep(2)


def xs_get_latest_qubedrive_date(folder):
    """Baca folder Updated, ambil tanggal MODIFIKASI TERBARU dari semua file
    di dalamnya (tidak bergantung pola nama file apa pun - fleksibel untuk
    format nama file apa saja). Return: datetime.date terbaru, atau None
    kalau folder kosong / tidak ditemukan / semua file gagal dibaca."""
    try:
        names = os.listdir(folder)
    except Exception as e:
        log.warning(f"  [XShelf][QubeDrive Freshness] Gagal baca folder '{folder}': {e}")
        return None

    latest = None
    for name in names:
        full_path = os.path.join(folder, name)
        try:
            if not os.path.isfile(full_path):
                continue
            mtime = os.path.getmtime(full_path)
        except Exception:
            continue
        d = datetime.fromtimestamp(mtime).date()
        if latest is None or d > latest:
            latest = d
    return latest


def xs_wait_for_fresh_qubedrive_data() -> bool:
    """Setelah download_data QubeDrive, tunggu (poll tiap
    QUBEDRIVE_FRESHNESS_POLL_INTERVAL detik, maksimal
    QUBEDRIVE_FRESHNESS_WAIT_MAX detik / 3-5 menit) sampai ada file baru
    bertanggal HARI INI di folder Updated. Kalau sampai batas waktu tetap
    data kemarin/lama, return False (skip lanjut ke EBarcode)."""
    today = datetime.now().date()
    log.info(f"[XShelf] STEP 0c - Cek folder Updated, tunggu data tanggal hari ini "
             f"({today.strftime('%d-%m-%Y')})...")

    start = time.time()
    latest = xs_get_latest_qubedrive_date(QUBEDRIVE_UPDATED_FOLDER)
    while latest != today and (time.time() - start) < QUBEDRIVE_FRESHNESS_WAIT_MAX:
        elapsed = int(time.time() - start)
        info_tgl = latest.strftime('%d-%m-%Y') if latest else "tidak ditemukan"
        log.info(f"  ... belum ada data hari ini (data terbaru di folder: {info_tgl}), "
                 f"menunggu... ({elapsed}s/{QUBEDRIVE_FRESHNESS_WAIT_MAX}s)")
        time.sleep(QUBEDRIVE_FRESHNESS_POLL_INTERVAL)
        latest = xs_get_latest_qubedrive_date(QUBEDRIVE_UPDATED_FOLDER)

    if latest == today:
        log.info(f"  OK: Data terbaru ditemukan (tanggal {latest.strftime('%d-%m-%Y')})")
        return True

    info_tgl = latest.strftime('%d-%m-%Y') if latest else "tidak ada file yang cocok"
    log.warning(f"*** [XShelf][QubeDrive Freshness] Data di folder Updated BUKAN data hari ini "
                f"(terbaru: {info_tgl}). Proses dihentikan, tidak lanjut ke EBarcode.")
    return False


def xs_fail(msg: str) -> bool:
    log.error(f"*** [XShelf] FAIL: {msg}")
    xs_kill_ebarcode()
    return False


def xs_run_flow(run_qubedrive: bool = True, include_checkbox5: bool = True) -> bool:
    log.info("=== [XShelf] START EBarcode ===")

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if leftover:
        log.warning("  [XShelf] EBarcode masih terbuka dari sesi sebelumnya, tutup dulu...")
        xs_kill_ebarcode()
        time.sleep(0.75)
    else:
        log.info("  [XShelf] EBarcode tidak terbuka, lanjut")

    if run_qubedrive:
        if not xs_run_qubedrive_upload():
            return xs_fail("Gagal menjalankan QubeDrive.exe -Jupload_data")
        if not xs_run_qubedrive_download():
            return xs_fail("Gagal menjalankan QubeDrive.exe -Jdownload_data")
        if not xs_wait_for_fresh_qubedrive_data():
            jam = QUBEDRIVE_RETRY_WAIT_SECONDS / 3600
            log.warning(f"*** [XShelf] Data belum update. Tunggu {jam:.0f} jam, lalu coba lagi "
                        f"(download saja, TANPA upload)...")
            time.sleep(QUBEDRIVE_RETRY_WAIT_SECONDS)

            log.info("[XShelf] Retry setelah menunggu - jalankan download_data lagi (tanpa upload)")
            if not xs_run_qubedrive_download():
                return xs_fail("Gagal menjalankan QubeDrive.exe -Jdownload_data (retry)")
            if not xs_wait_for_fresh_qubedrive_data():
                return xs_fail(f"Data QubeDrive masih belum update setelah menunggu {jam:.0f} jam "
                                f"dan retry download - proses dihentikan (exit)")

    log.info("[XShelf] STEP 1 - Buka EBarcode")
    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        time.sleep(2)

    # CATATAN FIX: cek dulu APAKAH EBarcode SUDAH ada di TfrmPrint (misal
    # diklik manual sebelumnya, atau sisa proses yang belum ditutup). Kalau
    # sudah, LANGSUNG lanjut - lewati klik TButton3 dan lewati juga
    # menunggu TfrmPrint (karena sudah pasti ada, tidak perlu ditunggu).
    print_win = window_already_open("TfrmPrint")
    if print_win:
        log.info("  [XShelf] Terdeteksi EBarcode SUDAH di TfrmPrint (Print sudah "
                 "diklik sebelumnya) - lewati STEP 2 & STEP 2b, langsung lanjut")
    else:
        if not xs_wait_win("TfrmBarcode"):
            return xs_fail("TfrmBarcode tidak muncul")

        app = xs_get_app()
        frm = app.window(class_name="TfrmBarcode")
        xs_safe_set_focus(frm)
        time.sleep(0.35)
        log.info("  OK: TfrmBarcode aktif")

        log.info("[XShelf] STEP 2 - Klik Print (TButton3)")
        try:
            click_control_message(frm.child_window(class_name="TButton", found_index=2))
            time.sleep(0.5)
            log.info("  OK")
        except Exception as e:
            return xs_fail(f"Gagal klik TButton3: {e}")

        log.info("[XShelf] STEP 2b - Tunggu window Printing Module (TfrmPrint)")
        print_win = xs_wait_win("TfrmPrint")
        if not print_win:
            return xs_fail("Window TfrmPrint (Printing Module) tidak muncul")
        log.info("  OK: TfrmPrint muncul")

    log.info("[XShelf] STEP 2c - Klik Export SHELF (TcxButton6)")
    export_win = window_already_open("TdlgExportSHELF")
    if export_win:
        log.info("  [XShelf] Terdeteksi dialog TdlgExportSHELF SUDAH terbuka (Export "
                 "SHELF sudah diklik sebelumnya) - lewati klik Export SHELF, langsung lanjut")
    else:
        try:
            app = xs_get_app()
            print_frm = app.window(class_name="TfrmPrint")
            xs_safe_set_focus(print_frm)
            time.sleep(0.2)
            export_shelf_btn = print_frm.child_window(title="Export SHELF", class_name="TcxButton")
            if not export_shelf_btn.exists():
                return xs_fail("Tombol 'Export SHELF' tidak ditemukan di TfrmPrint")

            # CATATAN FIX: "Export SHELF" adalah TcxButton (DevExpress/skinned,
            # custom-drawn). Tombol jenis ini kadang TIDAK bereaksi terhadap
            # BM_CLICK walau SendMessage-nya "sukses" tanpa error - akibatnya
            # click_control_message() lama langsung return True (dianggap
            # berhasil) padahal di layar TIDAK terjadi apa-apa, dan tampilan
            # jadi macet/stuck di TfrmPrint karena dialog TdlgExportSHELF
            # tidak pernah muncul. Di sini klik DIVERIFIKASI: kalau dialog
            # belum muncul, otomatis susulkan simulasi mouse asli
            # (WM_LBUTTONDOWN/UP lewat PostMessage - kursor fisik TETAP tidak
            # ikut bergerak), diulang beberapa kali sebelum benar2 gagal.
            clicked_ok = False
            for attempt in range(1, 4):
                log.info(f"    [Export SHELF] Percobaan klik ke-{attempt}...")
                click_control_message(export_shelf_btn)
                time.sleep(0.6)
                export_win = window_already_open("TdlgExportSHELF")
                if export_win:
                    clicked_ok = True
                    break

                log.warning(f"    [Export SHELF] BM_CLICK tidak bereaksi, susulkan simulasi mouse asli...")
                try:
                    app = xs_get_app()
                    print_frm = app.window(class_name="TfrmPrint")
                    export_shelf_btn = print_frm.child_window(title="Export SHELF", class_name="TcxButton")
                    hwnd = export_shelf_btn.handle
                    rect = export_shelf_btn.rectangle()
                    cx = max((rect.right - rect.left) // 2, 1)
                    cy = max((rect.bottom - rect.top) // 2, 1)
                    lparam = (cy << 16) | (cx & 0xFFFF)
                    ctypes.windll.user32.PostMessageW(hwnd, 0x0200, 0, lparam)  # WM_MOUSEMOVE
                    time.sleep(0.05)
                    ctypes.windll.user32.PostMessageW(hwnd, _WM_LBUTTONDOWN, 1, lparam)
                    time.sleep(0.08)
                    ctypes.windll.user32.PostMessageW(hwnd, _WM_LBUTTONUP, 0, lparam)
                except Exception as e:
                    log.warning(f"    [Export SHELF] Simulasi mouse gagal: {e}")

                time.sleep(0.8)
                export_win = window_already_open("TdlgExportSHELF")
                if export_win:
                    clicked_ok = True
                    break

                log.warning(f"    [Export SHELF] Belum berhasil (percobaan {attempt}/3), coba lagi...")
                time.sleep(0.5)

            if not clicked_ok:
                return xs_fail("Gagal klik Export SHELF setelah beberapa percobaan "
                                "(dialog TdlgExportSHELF tidak kunjung muncul - tampilan stuck di TfrmPrint)")
            log.info("  OK: Export SHELF diklik, dialog TdlgExportSHELF terkonfirmasi muncul")
        except Exception as e:
            return xs_fail(f"Gagal klik Export SHELF: {e}")

    log.info("[XShelf] STEP 3 - Dialog Export | StoreSHELF")
    if not export_win:
        export_win = window_already_open("TdlgExportSHELF")
    if not export_win:
        return xs_fail("Dialog TdlgExportSHELF (Export | StoreSHELF) tidak muncul")
    log.info("  OK: Dialog TdlgExportSHELF siap dipakai")

    log.info("[XShelf] STEP 3b - Ceklis checkbox export")
    try:
        app = xs_get_app()
        dlg = app.window(class_name="TdlgExportSHELF")
        xs_safe_set_focus(dlg)
        time.sleep(0.2)

        checkbox_list = [("TCheckBox5", 4), ("TCheckBox3", 2), ("TCheckBox6", 5)]
        if not include_checkbox5:
            checkbox_list = [(n, i) for n, i in checkbox_list if n != "TCheckBox5"]
        log.info("  Checkbox yang akan diceklis: " + ", ".join(n for n, _ in checkbox_list))

        for cb_name, cb_index in checkbox_list:
            chk = dlg.child_window(class_name="TCheckBox", found_index=cb_index)
            click_control_message(chk)
            time.sleep(0.2)
            log.info(f"  OK: {cb_name} diceklis")
    except Exception as e:
        return xs_fail(f"Gagal ceklis checkbox: {e}")

    log.info("[XShelf] STEP 3c - Klik Search (TcxButton4)")
    try:
        app = xs_get_app()
        dlg = app.window(class_name="TdlgExportSHELF")
        xs_safe_set_focus(dlg)
        time.sleep(0.2)
        click_control_message(dlg.child_window(class_name="TcxButton", found_index=3))
        time.sleep(0.5)
        log.info("  OK: Search diklik")
    except Exception as e:
        return xs_fail(f"Gagal klik TcxButton4 (Search): {e}")

    log.info(f"[XShelf] STEP 4 - Menunggu popup Qube Barcode Module (Done), "
             f"maksimal {WAIT_TIMER // 60} menit...")
    start_wait = time.time()
    last_logged_min = -1
    popup = None
    while time.time() - start_wait < WAIT_TIMER:
        wins = Desktop(backend="win32").windows(class_name="TMessageForm")
        if wins:
            popup = wins[0]
            break
        elapsed_min = int((time.time() - start_wait) // 60)
        if elapsed_min != last_logged_min:
            remaining_min = (WAIT_TIMER // 60) - elapsed_min
            log.info(f"  ... belum ada popup, menunggu (sisa maks ~{remaining_min} menit)")
            last_logged_min = elapsed_min
        time.sleep(2)

    if popup is None:
        return xs_fail(f"Popup TMessageForm (Done) tidak muncul dalam {WAIT_TIMER // 60} menit")

    log.info(f"  OK: Popup muncul setelah ~{int(time.time() - start_wait)} detik, langsung diklik")
    try:
        app = xs_get_app()
        msg_dlg = app.window(class_name="TMessageForm")
        xs_safe_set_focus(msg_dlg)
        time.sleep(0.2)
        click_control_message(msg_dlg.child_window(class_name="TButton", found_index=0))
        time.sleep(0.3)
        log.info("  OK: Popup Done ditutup")
    except Exception as e:
        return xs_fail(f"Gagal klik TButton1 (Done): {e}")

    log.info("[XShelf] STEP 6 - Klik Z-Send (TcxButton3)")
    try:
        app = xs_get_app()
        dlg = app.window(class_name="TdlgExportSHELF")
        xs_safe_set_focus(dlg)
        time.sleep(0.2)
        click_control_message(dlg.child_window(class_name="TcxButton", found_index=2))
        time.sleep(0.5)
        log.info("  OK: Z-Send diklik")
    except Exception as e:
        return xs_fail(f"Gagal klik TcxButton3 (Z-Send): {e}")

    log.info("[XShelf] STEP 7 - Tutup EBarcode")
    xs_kill_ebarcode()

    log.info("=== [XShelf] SELESAI ===")
    return True


def stage_xshelf() -> bool:
    log.info("### TAHAP 1 - XShelf (RUN #1 & RUN #2, logika ORIGINAL) ###")
    log.info("### RUN #1 (dengan TCheckBox5) ###")
    ok1 = xs_run_flow(run_qubedrive=True, include_checkbox5=True)
    if ok1:
        log.info("RUN #1 selesai dengan sukses.")
    else:
        log.warning("RUN #1 selesai dengan error.")

    log.info("### RUN #2 (tanpa TCheckBox5) ###")
    ok2 = xs_run_flow(run_qubedrive=False, include_checkbox5=False)
    if ok2:
        log.info("RUN #2 selesai dengan sukses.")
    else:
        log.warning("RUN #2 selesai dengan error.")

    if ok1 and ok2:
        log.info("### TAHAP 1 (XShelf) - SEMUA PROSES SUKSES ###")
        return True
    log.warning("### TAHAP 1 (XShelf) - ADA PROSES YANG ERROR ###")
    return False


# ============================================================================
# ==========================  TAHAP 2: CHECKDISC  ===========================
# (headless — turunan dari CheckDisc.py, tanpa GUI/notifikasi, prefix cd_)
# ============================================================================
CD_ROW_MID_PATTERN = re.compile(
    r',,,(\d+\.\d{2}),(\d{2}/\d{2}/\d{4}),(\d{2}/\d{2}/\d{4}),(\d+\.\d{2}),'
    r'(\d{2}/\d{2}/\d{4}),(\d{2}/\d{2}/\d{4}),(\d+),,(\d{2}/\d{2}/\d{4}),,'
)
CD_SKU_START_PATTERN = re.compile(r'^\d+,')
CD_MAX_LINE_MERGE = 5


def cd_parse_date(date_str):
    if not date_str or date_str.strip() == "":
        return None
    try:
        return datetime.strptime(date_str.strip(), DATE_FORMAT)
    except ValueError:
        for fmt in ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except ValueError:
                continue
        return None


def cd_get_balance(value):
    try:
        v = value.strip().replace(",", ".")
        return float(v) if v else 0.0
    except Exception:
        return 0.0


def cd_is_7digit(val):
    v = val.strip()
    return len(v) == 7 and v.isdigit()


def cd_is_shelf_kosong(shelf_val):
    v = (shelf_val or "").strip()
    if v == "" or v == "-" or v == "0":
        return True
    if v.isdigit():
        return True
    return False


def cd_repair_broken_row(line):
    m = CD_ROW_MID_PATTERN.search(line)
    if not m:
        return None
    left = line[:m.start()]
    right = line[m.end():]
    left_parts = left.split(",")
    if len(left_parts) < 9:
        return None
    sku_f, shelf_f, price_f = left_parts[0], left_parts[1], left_parts[2]
    category6 = left_parts[-6:]
    desc_fragments = left_parts[3:-6]
    full_desc = ",".join(desc_fragments)
    desc_fixed = full_desc.replace(",", ";")
    field23 = right.split(",")[0] if right else sku_f
    desc_trunc = desc_fixed[:20]
    return [sku_f, shelf_f, price_f, desc_fixed] + category6 + [
        "", "", m.group(1), m.group(2), m.group(3), m.group(4),
        m.group(5), m.group(6), m.group(7), "", m.group(8), "",
        field23, desc_trunc,
    ]


def cd_merge_broken_lines(lines):
    merged = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if not line:
            i += 1
            continue
        current = line
        j = i
        merges_here = 0
        while (len(current.split(",")) < 23 and j + 1 < n
               and lines[j + 1] and not CD_SKU_START_PATTERN.match(lines[j + 1])
               and merges_here < CD_MAX_LINE_MERGE):
            current = current + lines[j + 1]
            j += 1
            merges_here += 1
        merged.append((i + 1, current, merges_here))
        i = j + 1
    return merged


def cd_baca_data_txt(filename):
    raw_list = []
    repaired_count = 0
    merged_count = 0
    skipped_count = 0
    if not os.path.exists(filename):
        log.error(f"[CheckDisc] File input tidak ditemukan: {filename}")
        return []
    with open(filename, "r", encoding="utf-8", errors="ignore", newline="") as f:
        raw_text = f.read()
    lineend = "\r\n" if "\r\n" in raw_text else "\n"
    raw_lines = [l.rstrip("\r\n") for l in raw_text.split(lineend)]

    for row_num, line, merges_here in cd_merge_broken_lines(raw_lines):
        line = line.strip()
        if not line:
            continue
        if merges_here > 0:
            merged_count += 1
        row = line.split(",")
        if len(row) > 24:
            repaired = cd_repair_broken_row(line)
            if repaired is not None:
                row = repaired
                repaired_count += 1
            else:
                skipped_count += 1
                continue
        if len(row) < 23:
            skipped_count += 1
            continue
        raw_list.append({
            "sku": row[0].strip(),
            "shelf": row[1].strip(),
            "price": row[2].strip(),
            "desc": row[3].strip(),
            "discount": row[12].strip(),
            "start_date_str": row[13].strip(),
            "end_date_str": row[14].strip(),
            "system_balance": row[18].strip(),
            "balance_num": cd_get_balance(row[18]),
            "start_date": cd_parse_date(row[13].strip()),
            "end_date": cd_parse_date(row[14].strip()),
            "col23": row[22].strip() if len(row) > 22 else "",
            "row_num": row_num,
        })
    log.info(f"[CheckDisc] Baca {len(raw_list)} baris valid (repaired={repaired_count}, "
             f"merged={merged_count}, skipped={skipped_count})")
    return raw_list


def cd_deduplicate_by_sku(raw_list):
    """Dedup berdasarkan SKU (kolom nomor 1), BUKAN deskripsi - kalau ada
    baris dengan SKU yang sama persis (duplikat baris untuk SKU yang sama
    di file input), cukup diambil satu (baris pertama yang ditemukan)."""
    sku_groups = defaultdict(list)
    for item in raw_list:
        sku_groups[item["sku"]].append(item)
    kept = []
    removed = []
    for sku, items in sku_groups.items():
        kept.append(items[0])
        removed.extend(items[1:])
    return kept, removed


def cd_parse_price(price_str):
    try:
        return round(float(str(price_str).strip().replace(",", ".")), 2)
    except Exception:
        return 0.0


def cd_price_history_path():
    return os.path.join(BASE_DIR, PRICE_HISTORY_FILE)


def cd_load_price_history():
    path = cd_price_history_path()
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                fixed = {}
                for sku, val in data.items():
                    if isinstance(val, dict):
                        fixed[sku] = val
                    else:
                        fixed[sku] = {"price": val, "desc": ""}
                return fixed
    except Exception as e:
        log.warning(f"[CheckDisc] Gagal load price history: {e}")
    return {}


def cd_save_price_history(history_map):
    path = cd_price_history_path()
    try:
        os.makedirs(BASE_DIR, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(history_map, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log.warning(f"[CheckDisc] Gagal simpan price history: {e}")


def cd_check_data_changes(current_items):
    old_history = cd_load_price_history()
    increases, decreases, desc_changes = [], [], []
    new_history = {}
    for item in current_items:
        sku = item["sku"]
        new_price = cd_parse_price(item["price"])
        new_desc = item.get("desc", "")
        new_history[sku] = {"price": new_price, "desc": new_desc}

        old_entry = old_history.get(sku)
        if old_entry is None:
            continue

        old_price = old_entry.get("price")
        old_desc = old_entry.get("desc", "")
        eligible = item.get("balance_num", 0) > 0

        if eligible and old_price is not None and new_price != old_price:
            changed = dict(item)
            changed["old_price"] = old_price
            changed["new_price"] = new_price
            changed["price_diff"] = new_price - old_price
            (increases if new_price > old_price else decreases).append(changed)

        if eligible and old_desc and new_desc != old_desc:
            changed_desc = dict(item)
            changed_desc["old_desc"] = old_desc
            changed_desc["new_desc"] = new_desc
            desc_changes.append(changed_desc)

    cd_save_price_history(new_history)
    return increases, decreases, desc_changes


_sku_order_cache = None


def cd_load_sku_order(path=None):
    """Baca file urutan SKU per shelf (CSV/TXT: Shelf,SKU,NomorUrut).
    Hasil di-cache supaya file cukup dibaca sekali per proses.
    Baris dengan NomorUrut kosong / tidak valid dilewati (SKU itu berarti
    tanpa filter urutan, jalan seperti biasa)."""
    global _sku_order_cache
    if _sku_order_cache is not None:
        return _sku_order_cache

    path = path or SKU_ORDER_FILE
    order_map = {}
    if not os.path.exists(path):
        log.info(f"[CheckDisc] File urutan SKU '{path}' tidak ada, lewati (semua item tanpa filter urutan).")
        _sku_order_cache = order_map
        return order_map

    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                if not row or len(row) < 2:
                    continue
                shelf = row[0].strip().lower()
                sku = row[1].strip()
                if not sku or sku.lower() == "sku":   # lewati kalau baris header
                    continue
                order_raw = row[2].strip() if len(row) >= 3 else ""
                if order_raw == "":
                    continue   # kosong = tanpa filter urutan
                try:
                    order_val = float(order_raw)
                except ValueError:
                    continue
                order_map[(shelf, sku)] = order_val
    except Exception as e:
        log.warning(f"[CheckDisc] Gagal baca file urutan SKU '{path}': {e}")

    log.info(f"[CheckDisc] Load urutan SKU dari '{path}': {len(order_map)} SKU punya nomor urut.")
    _sku_order_cache = order_map
    return order_map


def cd_sort_by_shelf(data):
    """Urutkan berdasarkan shelf, lalu (kalau ada) berdasarkan nomor urut
    SKU dari SKU_ORDER_FILE. Item yang tidak punya nomor urut (kosong /
    tidak terdaftar) ditaruh paling akhir di shelf-nya, urutan aslinya
    tetap dipertahankan (sort stabil) - berarti tanpa filter urutan sama
    sekali kalau file SKU_ORDER_FILE tidak ada / kosong."""
    order_map = cd_load_sku_order()

    def sort_key(item):
        shelf = item.get("shelf", "").strip().lower()
        sku = item.get("sku", "").strip()
        order = order_map.get((shelf, sku))
        return (shelf, order if order is not None else float("inf"))

    return sorted(data, key=sort_key)


def cd_dated_filename(base_name):
    date_str = datetime.now().strftime("%d-%m-%Y")
    return f"{base_name} ({date_str}).txt"


def cd_auto_save(base_name, lines):
    """Simpan file export dengan nama '<base_name> (tgl-hari-ini).txt'.
    Sebelum menulis, hapus dulu file lama dengan base_name yang sama
    (tanggal berapa pun) supaya di folder output HANYA ada 1 file per
    jenis (tidak menumpuk tiap hari — otomatis di-replace)."""
    os.makedirs(CHECKDISC_OUTPUT_PATH, exist_ok=True)

    pattern = os.path.join(CHECKDISC_OUTPUT_PATH, f"{base_name} (*).txt")
    for old_path in glob.glob(pattern):
        try:
            os.remove(old_path)
            log.info(f"[CheckDisc] Hapus file lama: {old_path}")
        except Exception as e:
            log.warning(f"[CheckDisc] Gagal hapus file lama {old_path}: {e}")

    filename = cd_dated_filename(base_name)
    path = os.path.join(CHECKDISC_OUTPUT_PATH, filename)
    with open(path, "w", encoding="utf-8") as f:
        for line in lines:
            f.write(line + chr(10))
    return path


def cd_export_price_change(price_increases, price_decreases):
    combined = list(price_increases) + list(price_decreases)
    if not combined:
        return None
    data = cd_sort_by_shelf(combined)
    lines = [item["sku"] + ",1" for item in data]
    path = cd_auto_save("PRICE CHANGE", lines)
    log.info(f"[CheckDisc] Export PRICE CHANGE: {len(data)} item -> {path}")
    return path


def cd_export_desc_change(desc_changes):
    if not desc_changes:
        return None
    data = cd_sort_by_shelf(desc_changes)
    lines = [item["sku"] + ",1" for item in data]
    path = cd_auto_save("CHANGE DESCRIPTION", lines)
    log.info(f"[CheckDisc] Export CHANGE DESCRIPTION: {len(data)} item -> {path}")
    return path


def cd_export_shelf_kosong(dedup_data):
    data = [d for d in dedup_data
            if d.get("balance_num", 0) > 0 and cd_is_shelf_kosong(d.get("shelf", ""))]
    if not data:
        return None
    data = cd_sort_by_shelf(data)
    lines = [item["sku"] + ",1" for item in data]
    path = cd_auto_save("NO SHELF", lines)
    log.info(f"[CheckDisc] Export NO SHELF: {len(data)} item -> {path}")
    return path


def cd_export_qty_minus(dedup_data):
    """Export item yang qty system-nya MINUS (balance_num < 0) ke
    'QTY MINUS (tanggal).txt'. Tidak difilter balance>0 karena justru
    yang dicari adalah balance NEGATIF."""
    data = [d for d in dedup_data if d.get("balance_num", 0) < 0]
    if not data:
        return None
    data = cd_sort_by_shelf(data)
    lines = []
    for item in data:
        qty = abs(item.get("balance_num", 0))
        qty_str = str(int(qty)) if qty == int(qty) else str(qty)
        lines.append(f"{item['sku']},{qty_str}")
    path = cd_auto_save(QTY_MINUS_BASENAME, lines)
    log.info(f"[CheckDisc] Export QTY MINUS: {len(data)} item -> {path}")
    return path


def cd_compute_berjalan_expired(dedup_data):
    """Pisahkan item balance>0 menjadi 'berjalan' (promo masih aktif / tanpa
    diskon) dan 'expired' (promo sudah lewat tanggal end_date). Sama seperti
    proses_data() di CheckDisc.py asli."""
    filtered_data = [d for d in dedup_data if d.get("balance_num", 0) > 0]
    today = datetime.now()
    berjalan, expired = [], []
    for item in filtered_data:
        end_date = item.get("end_date")
        if end_date is None:
            expired.append(item)
        else:
            end_date_only = end_date.replace(hour=23, minute=59, second=59)
            if end_date_only >= today:
                berjalan.append(item)
            else:
                expired.append(item)
    return berjalan, expired


def cd_get_recent_start_items(dedup_data, lookback_days=START_LOOKBACK_DAYS):
    """Item balance>0 yang promo-nya BARU MULAI: start_date antara
    (hari ini - lookback_days) sampai hari ini (inklusif kedua sisi)."""
    today = datetime.now()
    today_only = today.replace(hour=23, minute=59, second=59, microsecond=0)
    result = []
    for item in dedup_data:
        if item.get("balance_num", 0) <= 0:
            continue
        start_date = item.get("start_date")
        if start_date is None:
            continue
        if start_date <= today_only and (today - start_date) < timedelta(days=lookback_days):
            result.append(item)
    return result


def cd_export_price_adjustment(expired_items, recent_start_items=None,
                                lookback_days=EXPIRED_LOOKBACK_DAYS):
    """BEST PRICE / 'promo sudah tidak berjalan' check — turunan dari
    show_expired_alert() di CheckDisc.py asli. Mengambil item yang
    end_date-nya sudah lewat DAN selisihnya dari hari ini <= lookback_days
    (semula 30 hari / ~2 minggu, sekarang di-set 3 hari).

    Digabung juga dengan item yang promo-nya BARU MULAI (start_date = hari
    ini atau dalam START_LOOKBACK_DAYS hari terakhir), lalu semuanya
    di-export ke PRICE ADJUSTMENT (tanggal).txt, dedup by SKU."""
    today = datetime.now()
    recent_expired = [
        item for item in (expired_items or [])
        if item.get("end_date") is not None
        and item["end_date"].replace(hour=23, minute=59, second=59) < today
        and (today - item["end_date"].replace(hour=23, minute=59, second=59)) < timedelta(days=lookback_days)
    ]

    combined = {}
    for item in recent_expired:
        combined[item["sku"]] = item
    for item in (recent_start_items or []):
        combined.setdefault(item["sku"], item)

    if not combined:
        return None
    data = cd_sort_by_shelf(list(combined.values()))
    lines = [item["sku"] + ",1" for item in data]
    path = cd_auto_save("PRICE ADJUSTMENT", lines)
    log.info(f"[CheckDisc] Export PRICE ADJUSTMENT (promo baru berakhir <= {lookback_days} hari, "
             f"atau baru mulai <= {START_LOOKBACK_DAYS} hari): {len(data)} item -> {path}")
    return path


def stage_checkdisc(input_path=CHECKDISC_INPUT_PATH):
    """Return (ok: bool, ada_perubahan: bool)."""
    log.info("### TAHAP 2 - CheckDisc (export headless) ###")
    log.info(f"  Input  : {input_path}")
    log.info(f"  Output : {CHECKDISC_OUTPUT_PATH}")

    if not os.path.exists(input_path):
        log.error(f"*** [CheckDisc] FAIL: File input tidak ditemukan: {input_path}")
        return False, False

    raw_data = cd_baca_data_txt(input_path)
    if not raw_data:
        log.error("*** [CheckDisc] FAIL: Tidak ada data valid yang bisa dibaca.")
        return False, False

    dedup_data, removed = cd_deduplicate_by_sku(raw_data)
    log.info(f"[CheckDisc] Dedup: {len(dedup_data)} data, {len(removed)} dibuang")

    berjalan, expired = cd_compute_berjalan_expired(dedup_data)
    log.info(f"[CheckDisc] Aktif (berjalan): {len(berjalan)} | Sudah lewat (expired): {len(expired)}")

    price_increases, price_decreases, desc_changes = cd_check_data_changes(dedup_data)
    log.info(f"[CheckDisc] Price naik: {len(price_increases)} | Price turun: {len(price_decreases)} | "
             f"Deskripsi berubah: {len(desc_changes)}")

    exported = []

    # BEST PRICE / cek item yang promo-nya baru selesai (<= EXPIRED_LOOKBACK_DAYS hari)
    # ATAU baru mulai (<= START_LOOKBACK_DAYS hari) -> gabung ke PRICE ADJUSTMENT
    recent_start_items = cd_get_recent_start_items(dedup_data)
    log.info(f"[CheckDisc] Promo baru mulai (<= {START_LOOKBACK_DAYS} hari): {len(recent_start_items)}")
    if expired or recent_start_items:
        p = cd_export_price_adjustment(expired, recent_start_items)
        if p:
            exported.append(p)

    if price_increases or price_decreases:
        p = cd_export_price_change(price_increases, price_decreases)
        if p:
            exported.append(p)
    if desc_changes:
        p = cd_export_desc_change(desc_changes)
        if p:
            exported.append(p)
    p = cd_export_shelf_kosong(dedup_data)
    if p:
        exported.append(p)
    p = cd_export_qty_minus(dedup_data)
    if p:
        exported.append(p)

    if not exported:
        log.info("[CheckDisc] Tidak ada perubahan terdeteksi — tidak ada file yang di-export.")
        return True, False

    log.info(f"[CheckDisc] Selesai. Total file di-export: {len(exported)}")
    for p in exported:
        log.info(f"  -> {p}")
    return True, True


# ============================================================================
# =========================  TAHAP 3: PRINTAUTO24  ===========================
# (trimmed — hanya Step 1,2,3(Zebex),4(TdlgData), prefix pa_)
# ============================================================================
# Aturan print (PrintAuto24 + QTY MINUS):
#   - Kalau HARI INI ada file PRICE ADJUSTMENT / PRICE CHANGE / CHANGE DESCRIPTION
#     -> HANYA file price itu yang di-print. NO SHELF dan QTY MINUS cukup
#        disimpan di folder output (tidak dicentang, tidak di-print).
#   - Kalau TIDAK ada satupun file price hari ini
#     -> NO SHELF (PrintAuto24) dan QTY MINUS (TAHAP 4) di-print.
PA_PRICE_BASENAMES = [
    "PRICE ADJUSTMENT",
    "PRICE CHANGE",
    "CHANGE DESCRIPTION",
]
PA_SHELF_BASENAMES = [
    "NO SHELF",
]
PA_ALLOWED_BASENAMES = PA_PRICE_BASENAMES + PA_SHELF_BASENAMES


def _pa_today_filenames(basenames):
    date_str = datetime.now().strftime("%d-%m-%Y")
    return [f"{base} ({date_str}).TXT".upper() for base in basenames]


def _pa_existing_today(basenames):
    """Nama file (UPPER) milik basenames yang BENAR-BENAR ada di folder output hari ini."""
    try:
        on_disk = {n.upper() for n in os.listdir(CHECKDISC_OUTPUT_PATH)}
    except Exception:
        on_disk = set()
    return [n for n in _pa_today_filenames(basenames) if n in on_disk]


def pa_price_data_today():
    """File price (PRICE ADJUSTMENT/PRICE CHANGE/CHANGE DESCRIPTION) yang ada hari ini."""
    return _pa_existing_today(PA_PRICE_BASENAMES)


def pa_today_allowed_items():
    """Daftar file yang boleh dicentang/di-print di TdlgData (PrintAuto24).
    Ada data price -> hanya file price. Tidak ada -> NO SHELF."""
    price = pa_price_data_today()
    if price:
        log.info(f"  [PrintAuto24] Ada data price hari ini {price} -> "
                 f"NO SHELF TIDAK di-print (cukup tersimpan di output)")
        return set(price)

    shelf = _pa_existing_today(PA_SHELF_BASENAMES)
    if shelf:
        log.info(f"  [PrintAuto24] Tidak ada data price hari ini -> print NO SHELF {shelf}")
    else:
        log.info("  [PrintAuto24] Tidak ada data price maupun NO SHELF hari ini")
    return set(shelf)


def pa_wait_win(class_name, timeout=TIMEOUT):
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="win32").windows(class_name=class_name)
        if wins:
            return wins[0]
        time.sleep(0.25)
    return None


def pa_get_app():
    return Application(backend="win32").connect(path=EXE_PATH)


def pa_safe_set_focus(control):
    try:
        control.set_focus()
        return True
    except Exception as e:
        log.warning(f"[PrintAuto24] set_focus dilewati: {e}")
        return False


def pa_kill_ebarcode():
    log.warning("  [PrintAuto24][KILL] Menutup EBarcode...")
    for cls in ["TdlgData", "TdlgExportSHELF", "TdlgStoreSHELF", "TMessageForm", "TfrmPrint", "TfrmBarcode"]:
        wins = Desktop(backend="win32").windows(class_name=cls)
        for w in wins:
            try:
                pa_safe_set_focus(w)
                w.close()
                time.sleep(0.15)
            except Exception as e:
                log.warning(f"  [PrintAuto24][KILL] Gagal tutup {cls}: {e}")

    time.sleep(0.75)
    wins = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if wins:
        log.warning("  [PrintAuto24][KILL] EBarcode masih ada, paksa taskkill...")
        try:
            subprocess.run(["taskkill", "/F", "/IM", "EBarcode.exe"], capture_output=True, timeout=5)
        except Exception as e:
            log.warning(f"  [PrintAuto24][KILL] taskkill error: {e}")
    else:
        log.warning("  [PrintAuto24][KILL] EBarcode sudah tertutup")


def find_and_click_zebex(frm_print) -> bool:
    """Cari tombol Zebex di TfrmPrint lalu klik. Ada 3 cara dicoba
    berurutan, dan (BEDA dari versi lama) benar-benar dicek pakai
    .exists() - bukan cuma cek None, karena child_window() di pywinauto
    bersifat lazy dan TIDAK PERNAH None/exception walau kontrolnya
    sebenarnya tidak ada, sehingga fallback lama tidak pernah kepakai.
    Return True hanya kalau klik benar-benar terkirim ke control yang valid."""
    zebex = None

    # Cara 1: cari lewat title "Zebex"
    try:
        candidate = frm_print.child_window(title="Zebex", class_name="TcxButton")
        if candidate.exists():
            zebex = candidate
    except Exception:
        pass

    # Cara 2: scan semua TcxButton, cari yang teksnya mengandung "ZEBEX"
    if zebex is None:
        try:
            for btn in frm_print.children(class_name="TcxButton"):
                try:
                    if "ZEBEX" in btn.window_text().strip().upper():
                        zebex = btn
                        break
                except Exception:
                    pass
        except Exception:
            pass

    # Cara 3: fallback pakai found_index=8 (posisi tombol Zebex biasanya)
    if zebex is None:
        try:
            candidate = frm_print.child_window(class_name="TcxButton", found_index=8)
            if candidate.exists():
                zebex = candidate
        except Exception:
            pass

    if zebex is None:
        log.error("  [Zebex] Tombol Zebex tidak ditemukan lewat title, scan teks, maupun found_index=8.")
        return False

    ok = click_control_message(zebex)
    if not ok:
        log.error("  [Zebex] Tombol Zebex ditemukan tapi GAGAL diklik (click_control_message gagal).")
        return False

    return True


# Window utilitas "Zebex Download Utility" (aplikasi terpisah) yang kadang
# muncul setelah Zebex diklik, SEBELUM TdlgData terbuka. Kalau muncul harus
# ditutup dulu; kalau tidak muncul, langsung lanjut ke TdlgData.
PA_ZEBEX_UTIL_TITLE_RE = r"(?i).*zebex.*"
PA_ZEBEX_UTIL_CLOSE_BTN_CLASS = "TTeButton"   # = "TTeButton1" -> found_index=0
PA_EBARCODE_CLASSES = {"TfrmBarcode", "TfrmPrint", "TdlgData", "TdlgExportSHELF",
                       "TdlgStoreSHELF", "TMessageForm"}


def pa_find_zebex_utility():
    """Cek SEKALI apakah window Zebex Download Utility sedang terbuka."""
    try:
        for w in Desktop(backend="win32").windows(title_re=PA_ZEBEX_UTIL_TITLE_RE,
                                                  visible_only=True):
            try:
                if w.friendly_class_name() in PA_EBARCODE_CLASSES or \
                        w.class_name() in PA_EBARCODE_CLASSES:
                    continue
                return w
            except Exception:
                continue
    except Exception:
        pass
    return None


def pa_close_zebex_utility(wait_appear=3.0) -> bool:
    """Setelah klik Zebex: kalau muncul window Zebex Download Utility,
    tutup (klik TTeButton1, fallback close biasa). Kalau tidak muncul,
    atau TdlgData sudah terbuka duluan, langsung return tanpa error.
    Return True kalau ada utility yang ditutup."""
    start = time.time()
    util = None
    while time.time() - start < wait_appear:
        util = pa_find_zebex_utility()
        if util:
            break
        if window_already_open("TdlgData"):
            log.info("  [Zebex] Tidak ada Zebex Download Utility, langsung ke TdlgData")
            return False
        time.sleep(0.2)

    if not util:
        log.info("  [Zebex] Tidak ada Zebex Download Utility, lanjut ke TdlgData")
        return False

    log.info(f"  [Zebex] Zebex Download Utility terdeteksi ('{util.window_text()}') - menutup...")
    closed_by_button = False
    try:
        btn = util.child_window(class_name=PA_ZEBEX_UTIL_CLOSE_BTN_CLASS, found_index=0)
        if btn.exists():
            closed_by_button = click_control_message(btn)
            if closed_by_button:
                log.info("  [Zebex] Klik TTeButton1 terkirim")
    except Exception as e:
        log.warning(f"  [Zebex] Klik TTeButton1 gagal: {e}")

    # Tunggu hilang; kalau masih ada -> close biasa
    t0 = time.time()
    while time.time() - t0 < 2.0:
        if not pa_find_zebex_utility():
            log.info("  [Zebex] Zebex Download Utility tertutup")
            return True
        time.sleep(0.2)

    try:
        log.info("  [Zebex] Pakai close biasa...")
        util.close()
    except Exception as e:
        log.warning(f"  [Zebex] Close biasa gagal: {e}")
    time.sleep(0.5)
    if pa_find_zebex_utility():
        log.warning("  [Zebex] Zebex Download Utility masih terbuka (lanjut tetap dicoba)")
    return True


def pa_fail(msg: str) -> bool:
    log.error(f"*** [PrintAuto24] FAIL: {msg}")
    pa_kill_ebarcode()
    return False


def stage_printauto24() -> bool:
    log.info("### TAHAP 3 - PrintAuto24 (trimmed: Step 1,2,3-Zebex,4-TdlgData) ###")

    if not pa_today_allowed_items():
        log.info("[PrintAuto24] Tidak ada file yang perlu di-print hari ini, lewati print.")
        return True

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if leftover:
        log.warning("  [PrintAuto24] EBarcode masih terbuka dari sesi sebelumnya, tutup dulu...")
        pa_kill_ebarcode()
        time.sleep(0.75)

    log.info("[PrintAuto24] STEP 1 - Buka EBarcode")
    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        time.sleep(2)

    # CATATAN FIX: cek dulu APAKAH EBarcode SUDAH di TfrmPrint. Kalau
    # sudah, LANGSUNG lanjut - lewati klik TButton3 dan lewati menunggu
    # TfrmPrint (sudah pasti ada, tidak perlu ditunggu lagi).
    print_win = window_already_open("TfrmPrint")
    if print_win:
        log.info("  [PrintAuto24] Terdeteksi EBarcode SUDAH di TfrmPrint (Print sudah "
                 "diklik sebelumnya) - lewati STEP 2, langsung lanjut")
    else:
        if not pa_wait_win("TfrmBarcode"):
            return pa_fail("TfrmBarcode tidak muncul")

        app = pa_get_app()
        frm = app.window(class_name="TfrmBarcode")
        pa_safe_set_focus(frm)
        time.sleep(0.35)
        log.info("  OK: TfrmBarcode aktif")

        log.info("[PrintAuto24] STEP 2 - Klik Print (TButton3)")
        try:
            click_control_message(frm.child_window(class_name="TButton", found_index=2))
            time.sleep(0.5)
            log.info("  OK")
        except Exception as e:
            return pa_fail(f"Gagal klik TButton3: {e}")

        print_win = pa_wait_win("TfrmPrint", timeout=15)
        if not print_win:
            return pa_fail("TfrmPrint tidak muncul")

    log.info("[PrintAuto24] STEP 3 - TfrmPrint -> klik Zebex")
    data_win = window_already_open("TdlgData")
    if data_win:
        log.info("  [PrintAuto24] Terdeteksi TdlgData SUDAH terbuka (Zebex sudah "
                 "diklik sebelumnya) - lewati klik Zebex, langsung lanjut")
    else:
        app = pa_get_app()
        frm_print = app.window(class_name="TfrmPrint")
        pa_safe_set_focus(frm_print)
        time.sleep(0.25)

        try:
            if not find_and_click_zebex(frm_print):
                return pa_fail("Gagal klik Zebex (lihat log [Zebex] di atas untuk detail)")
            time.sleep(0.5)
            log.info("  OK: Zebex diklik")
            pa_close_zebex_utility()
        except Exception as e:
            return pa_fail(f"Gagal klik Zebex: {e}")

    log.info("[PrintAuto24] STEP 4 - Dialog TdlgData")
    if not data_win:
        data_win = pa_wait_win("TdlgData", timeout=20)
        if not data_win:
            return pa_fail("TdlgData tidak muncul")

    app = pa_get_app()
    data_win = app.window(class_name="TdlgData")
    pa_safe_set_focus(data_win)
    time.sleep(0.3)

    allowed = pa_today_allowed_items()
    log.info(f"  Item yang boleh dicentang (hari ini): {sorted(allowed)}")

    try:
        click_control_message(data_win.child_window(class_name="TBitBtn", found_index=3))
        time.sleep(0.2)
        log.info("  Un-tick semua selesai")

        checklist = data_win.child_window(class_name="TCheckListBox", found_index=1)
        pa_safe_set_focus(checklist)
        time.sleep(0.15)

        hwnd = checklist.handle
        LB_GETCOUNT  = 0x018B
        LB_GETTEXT   = 0x0189
        LB_SETCURSEL = 0x0186
        LB_SETCHECK  = 0x0191
        LB_GETCHECK  = 0x0190
        LB_GETITEMHEIGHT = 0x01A1

        count = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCOUNT, 0, 0)
        log.info(f"  Jumlah item TCheckListBox: {count}")

        checked_any = False
        for i in range(count):
            buf = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.SendMessageW(hwnd, LB_GETTEXT, i, buf)
            item_text = buf.value.strip()
            log.info(f"    [{i}] '{item_text}'")

            if item_text.upper() in allowed:
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCURSEL, i, 0)
                time.sleep(0.075)
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCHECK, i, 1)
                time.sleep(0.1)
                checked = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCHECK, i, 0)
                if checked == 1:
                    log.info(f"    OK: dicentang -> '{item_text}'")
                    checked_any = True
                else:
                    log.warning(f"    LB_SETCHECK gagal untuk '{item_text}', coba klik area checkbox")
                    item_height = ctypes.windll.user32.SendMessageW(hwnd, LB_GETITEMHEIGHT, 0, 0)
                    if item_height <= 0:
                        item_height = 16
                    y = (i * item_height) + (item_height // 2)
                    ctypes.windll.user32.PostMessageW(hwnd, 0x0201, 1, (y << 16) | 8)
                    time.sleep(0.03)
                    ctypes.windll.user32.PostMessageW(hwnd, 0x0202, 0, (y << 16) | 8)
                    time.sleep(0.1)
                    checked_any = True
            else:
                log.info("      -> dilewati (bukan tanggal hari ini / bukan file target)")

        if not checked_any:
            log.warning("  Tidak ada item bertanggal hari ini di daftar — tidak ada yang dicentang.")
        time.sleep(0.15)
    except Exception as e:
        return pa_fail(f"Error saat proses checklist TdlgData: {e}")

    log.info("[PrintAuto24] STEP 4b - Extract >> (TBitBtn7)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=6))
        time.sleep(1.25)
        log.info("  OK: Extract selesai")
    except Exception as e:
        return pa_fail(f"Gagal klik Extract: {e}")

    log.info("[PrintAuto24] STEP 4c - Print >> (TBitBtn6)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=5))
        time.sleep(1.25)
        log.info("  OK: Print dikirim")
    except Exception as e:
        return pa_fail(f"Gagal klik Print: {e}")

    popup = pa_wait_win("TMessageForm", timeout=3600)
    if popup:
        try:
            app = pa_get_app()
            dlg = app.window(class_name="TMessageForm")
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            time.sleep(0.25)
            log.info("  OK: Popup print ditutup")
        except Exception as e:
            log.info(f"  Popup print error: {e}")
    else:
        log.info("  Tidak ada popup, lanjut")

    log.info("[PrintAuto24] STEP 4d - Close TdlgData (TBitBtn8)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=7))
        time.sleep(0.25)
        log.info("  OK: TdlgData ditutup")
    except Exception as e:
        log.info(f"  Close TdlgData error: {e}")

    pa_kill_ebarcode()
    log.info("=== [PrintAuto24] SELESAI ===")
    return True


# ============================================================================
# ==============  PRINT QTY MINUS via COLLECTING LABEL (menu list #4)  =======
# (adaptasi dari open_brown_box_print_form pada bb.py, label diganti jadi
#  COLLECTING LABEL, lalu lanjut proses print seperti output lain: Zebex ->
#  TdlgData -> centang hanya QTY MINUS (tanggal) -> Extract -> Print)
# ============================================================================
def pa_open_collecting_label_print_form():
    """
    Buka EBarcode (kalau belum jalan) -> di TfrmBarcode pilih 'COLLECTING
    LABEL' (menu list nomor 4) di Select Label -> klik Print (TButton3) ->
    tunggu TfrmPrint muncul.
    Return: frm_print (pywinauto window) kalau berhasil, None kalau gagal.
    """
    log.info("  [QTY MINUS] Buka EBarcode")

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    leftover_print = Desktop(backend="win32").windows(class_name="TfrmPrint")
    if leftover or leftover_print:
        log.warning("  [QTY MINUS] EBarcode/TfrmPrint masih terbuka dari sesi "
                    "sebelumnya, tutup dulu supaya tidak salah pilih menu...")
        pa_kill_ebarcode()
        time.sleep(0.75)

    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        time.sleep(1.5)

    if not pa_wait_win("TfrmBarcode", timeout=TIMEOUT):
        log.warning("  [QTY MINUS] TfrmBarcode tidak muncul, lewati print QTY MINUS")
        return None

    app = pa_get_app()
    frm = app.window(class_name="TfrmBarcode")
    pa_safe_set_focus(frm)
    time.sleep(0.25)

    log.info(f"  [QTY MINUS] Pilih label '{COLLECTING_LABEL_NAME}' di Select Label")
    try:
        combo = frm.child_window(class_name="TwwDBComboBox", found_index=0)
        pa_safe_set_focus(combo)
        time.sleep(0.1)

        WM_KEYDOWN = 0x0100
        WM_KEYUP   = 0x0101
        VK_HOME    = 0x24
        VK_DOWN    = 0x28
        VK_RETURN  = 0x0D
        KEYEVENTF_KEYUP = 0x0002

        def send_vkey(hwnd, vk):
            """FIX: dulu navigasi dikirim lewat PostMessage(WM_KEYDOWN/UP)
            langsung ke hwnd combo. Masalahnya, PostMessage BYPASS mekanisme
            keyboard-focus asli Windows - begitu dropdown combo terbuka,
            Windows biasanya memunculkan POPUP LISTBOX TERPISAH yang benar-
            benar memegang keyboard focus saat itu, BUKAN hwnd combo utama
            yang jadi target PostMessage. Akibatnya Home/Down yang di-post
            tidak pernah "dilihat" oleh listbox popup itu, dan Enter yang
            menyusul cuma meng-konfirmasi item yang sedang ke-highlight saat
            itu (defaultnya item pertama) -> makanya selalu balik ke menu
            pertama walau dropdown-nya sendiri sudah kebuka.

            Sekarang dipakai keybd_event (keystroke ASLI, level input OS,
            sama seperti click_input() untuk buka dropdown) - Windows akan
            mengirimkannya ke window MANAPUN yang benar-benar sedang pegang
            keyboard focus saat itu (yaitu popup listbox-nya), persis seperti
            kalau user menekan tombol panah di keyboard sungguhan. Parameter
            hwnd tidak lagi dipakai untuk kirim pesan, tapi dipertahankan
            supaya pemanggilnya tidak perlu diubah."""
            ctypes.windll.user32.keybd_event(vk, 0, 0, 0)
            time.sleep(0.03)
            ctypes.windll.user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

        def read_combo_text():
            try:
                return combo.window_text().strip()
            except Exception:
                return ""

        def open_dropdown():
            """Buka dropdown Select Label dengan klik SUNGGUHAN (safe click,
            menggerakkan kursor mouse asli), BUKAN klik lewat window message
            (BM_CLICK/PostMessage).

            CATATAN FIX: TBtnWinControl (tombol panah bawah combo) adalah
            custom control, bukan TButton standar - BM_CLICK yang dikirim
            lewat click_control_message() sering TIDAK memicu dropdown-nya
            benar-benar terbuka (fungsi tetap return True walau dropdown
            sebenarnya tidak terbuka). Akibatnya VK_HOME/VK_DOWN/VK_RETURN
            yang dikirim sesudahnya cuma jalan di combo yang TERTUTUP, jadi
            hasilnya sering nyangkut/balik ke item pertama (default) alih-
            alih pindah ke 'COLLECTING LABEL'.

            Makanya sekarang dropdown dibuka pakai klik asli (set_focus lalu
            click_input - klik mouse betulan ke tombol panah), supaya
            dropdown BENAR-BENAR terbuka dulu, baru navigasi keyboard
            (Home/Down/Enter) dikirim lewat PostMessage seperti sebelumnya."""
            dropdown_btn = frm.child_window(class_name="TBtnWinControl", found_index=0)
            try:
                pa_safe_set_focus(dropdown_btn)
                time.sleep(0.05)
                dropdown_btn.click_input()
            except Exception as e:
                log.warning(f"  [QTY MINUS] Klik asli dropdown gagal ({e}), coba fallback klik message...")
                try:
                    click_control_message(dropdown_btn)
                except Exception as e2:
                    log.warning(f"  [QTY MINUS] Gagal klik dropdown Select Label (fallback juga gagal): {e2}")
                    return False
            time.sleep(0.3)
            return True

        target_hwnd = combo.handle
        selected_ok = False

        MAX_ATTEMPTS = 3
        for attempt in range(1, MAX_ATTEMPTS + 1):
            if not open_dropdown():
                return None

            # PENTING: TIDAK set_focus() lagi ke combo di sini. Klik asli
            # (click_input) yang membuka dropdown di atas sudah otomatis
            # memberi keyboard focus yang BENAR ke popup listbox-nya.
            # Kalau di sini dipaksa set_focus() ke combo lagi, itu bisa
            # menutup/mengganggu popup listbox-nya (fokus asli tergeser
            # balik ke combo), sehingga keystroke asli di bawah (send_vkey)
            # malah tidak sampai ke listbox yang benar.

            # Reset ke item pertama (Home) lalu Down sebanyak
            # COLLECTING_LABEL_ITEM_INDEX kali (menu list nomor 4 -> index 3).
            send_vkey(target_hwnd, VK_HOME)
            time.sleep(0.1)
            for _ in range(COLLECTING_LABEL_ITEM_INDEX):
                send_vkey(target_hwnd, VK_DOWN)
                time.sleep(0.025)
            time.sleep(0.1)
            send_vkey(target_hwnd, VK_RETURN)
            time.sleep(0.25)

            result_text = read_combo_text()
            log.info(f"  [QTY MINUS] Percobaan {attempt}: Select Label -> '{result_text}'")

            if result_text.upper() == COLLECTING_LABEL_NAME.upper():
                selected_ok = True
                break
            log.warning(f"  [QTY MINUS] Percobaan {attempt} salah pilih "
                        f"('{result_text}' != '{COLLECTING_LABEL_NAME}'), akan dicoba ulang")

        if not selected_ok:
            log.warning(f"  [QTY MINUS] GAGAL set Select Label ke '{COLLECTING_LABEL_NAME}' "
                        f"setelah {MAX_ATTEMPTS}x percobaan, batalkan print QTY MINUS "
                        f"(supaya tidak mencetak label yang salah)")
            return None

        log.info(f"  OK: Select Label diset ke '{COLLECTING_LABEL_NAME}' (terverifikasi)")
    except Exception as e:
        log.warning(f"  [QTY MINUS] Gagal set Select Label COLLECTING LABEL: {e}")
        return None

    log.info("  [QTY MINUS] Klik Print (TButton3)")
    try:
        click_control_message(frm.child_window(class_name="TButton", found_index=2))
        time.sleep(0.5)
    except Exception as e:
        log.warning(f"  [QTY MINUS] Gagal klik TButton3: {e}")
        return None

    if not pa_wait_win("TfrmPrint"):
        log.warning("  [QTY MINUS] TfrmPrint tidak muncul")
        return None

    app = pa_get_app()
    frm_print = app.window(class_name="TfrmPrint")
    pa_safe_set_focus(frm_print)
    time.sleep(0.25)

    return frm_print


def pa_qty_minus_today_filename():
    date_str = datetime.now().strftime("%d-%m-%Y")
    return f"{QTY_MINUS_BASENAME} ({date_str}).txt"


def stage_print_qty_minus() -> bool:
    """Cetak file QTY MINUS (tanggal hari ini) lewat label COLLECTING LABEL,
    lalu lanjut proses print seperti output lain (Zebex -> TdlgData ->
    centang hanya QTY MINUS -> Extract -> Print)."""
    today_file = pa_qty_minus_today_filename()
    today_path = os.path.join(CHECKDISC_OUTPUT_PATH, today_file)
    if not os.path.exists(today_path):
        log.info(f"[QTY MINUS] Tidak ada file '{today_file}' hari ini, lewati print QTY MINUS.")
        return True

    price_today = pa_price_data_today()
    if price_today:
        log.info(f"[QTY MINUS] Ada data price hari ini {price_today} -> QTY MINUS TIDAK "
                 f"di-print (file '{today_file}' cukup tersimpan di output).")
        return True

    log.info("### TAHAP 4 - Print QTY MINUS (COLLECTING LABEL) ###")

    frm_print = pa_open_collecting_label_print_form()
    if frm_print is None:
        return pa_fail("Gagal membuka form print QTY MINUS (COLLECTING LABEL)")

    log.info("  [QTY MINUS] Klik Zebex")
    try:
        if not find_and_click_zebex(frm_print):
            return pa_fail("[QTY MINUS] Gagal klik Zebex (lihat log [Zebex] di atas untuk detail)")
        time.sleep(0.5)
        log.info("  OK: Zebex diklik")
        pa_close_zebex_utility()
    except Exception as e:
        return pa_fail(f"[QTY MINUS] Gagal klik Zebex: {e}")

    log.info("  [QTY MINUS] Dialog TdlgData")
    if not pa_wait_win("TdlgData", timeout=20):
        return pa_fail("[QTY MINUS] TdlgData tidak muncul")

    app = pa_get_app()
    data_win = app.window(class_name="TdlgData")
    pa_safe_set_focus(data_win)
    time.sleep(0.3)

    allowed = {today_file.upper()}
    log.info(f"  [QTY MINUS] Item yang boleh dicentang: {sorted(allowed)}")

    try:
        click_control_message(data_win.child_window(class_name="TBitBtn", found_index=3))
        time.sleep(0.2)
        log.info("  Un-tick semua selesai")

        checklist = data_win.child_window(class_name="TCheckListBox", found_index=1)
        pa_safe_set_focus(checklist)
        time.sleep(0.15)

        hwnd = checklist.handle
        LB_GETCOUNT  = 0x018B
        LB_GETTEXT   = 0x0189
        LB_SETCURSEL = 0x0186
        LB_SETCHECK  = 0x0191
        LB_GETCHECK  = 0x0190
        LB_GETITEMHEIGHT = 0x01A1

        count = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCOUNT, 0, 0)
        log.info(f"  Jumlah item TCheckListBox: {count}")

        checked_any = False
        for i in range(count):
            buf = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.SendMessageW(hwnd, LB_GETTEXT, i, buf)
            item_text = buf.value.strip()
            log.info(f"    [{i}] '{item_text}'")

            if item_text.upper() in allowed:
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCURSEL, i, 0)
                time.sleep(0.075)
                ctypes.windll.user32.SendMessageW(hwnd, LB_SETCHECK, i, 1)
                time.sleep(0.1)
                checked = ctypes.windll.user32.SendMessageW(hwnd, LB_GETCHECK, i, 0)
                if checked == 1:
                    log.info(f"    OK: dicentang -> '{item_text}'")
                    checked_any = True
                else:
                    log.warning(f"    LB_SETCHECK gagal untuk '{item_text}', coba klik area checkbox")
                    item_height = ctypes.windll.user32.SendMessageW(hwnd, LB_GETITEMHEIGHT, 0, 0)
                    if item_height <= 0:
                        item_height = 16
                    y = (i * item_height) + (item_height // 2)
                    ctypes.windll.user32.PostMessageW(hwnd, 0x0201, 1, (y << 16) | 8)
                    time.sleep(0.03)
                    ctypes.windll.user32.PostMessageW(hwnd, 0x0202, 0, (y << 16) | 8)
                    time.sleep(0.1)
                    checked_any = True
            else:
                log.info("      -> dilewati (bukan file QTY MINUS hari ini)")

        if not checked_any:
            log.warning("  [QTY MINUS] Tidak ada item QTY MINUS hari ini di daftar — tidak ada yang dicentang.")
        time.sleep(0.15)
    except Exception as e:
        return pa_fail(f"[QTY MINUS] Error saat proses checklist TdlgData: {e}")

    log.info("  [QTY MINUS] Extract >> (TBitBtn7)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=6))
        time.sleep(1.25)
        log.info("  OK: Extract selesai")
    except Exception as e:
        return pa_fail(f"[QTY MINUS] Gagal klik Extract: {e}")

    log.info("  [QTY MINUS] Print >> (TBitBtn6)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=5))
        time.sleep(1.25)
        log.info("  OK: Print dikirim")
    except Exception as e:
        return pa_fail(f"[QTY MINUS] Gagal klik Print: {e}")

    popup = pa_wait_win("TMessageForm", timeout=3600)
    if popup:
        try:
            app = pa_get_app()
            dlg = app.window(class_name="TMessageForm")
            click_control_message(dlg.child_window(class_name="TButton", found_index=0))
            time.sleep(0.25)
            log.info("  OK: Popup print ditutup")
        except Exception as e:
            log.info(f"  Popup print error: {e}")
    else:
        log.info("  Tidak ada popup, lanjut")

    log.info("  [QTY MINUS] Close TdlgData (TBitBtn8)")
    try:
        app = pa_get_app()
        click_control_message(app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=7))
        time.sleep(0.25)
        log.info("  OK: TdlgData ditutup")
    except Exception as e:
        log.info(f"  Close TdlgData error: {e}")

    pa_kill_ebarcode()
    log.info("=== [QTY MINUS] SELESAI ===")
    return True


# ============================================================================
# ==============================  ORKESTRATOR  ===============================
# ============================================================================
def gap(seconds=GAP_SECONDS):
    log.info(f"  ... jeda {seconds} detik sebelum lanjut ...")
    time.sleep(seconds)


def exit_with_countdown(seconds=EXIT_DELAY_SUCCESS):
    log.info(f"Semua proses selesai. Aplikasi akan tertutup otomatis dalam {seconds} detik...")
    time.sleep(seconds)
    log.info("Keluar.")


def close_leftover_apps_before_start():
    """Tutup dulu aplikasi QubeDrive.exe dan EBarcode.exe (kalau masih
    terbuka/berjalan dari sesi sebelumnya) sebelum memulai proses."""
    log.info("=== PRE-START - Tutup QubeDrive & EBarcode (jika masih berjalan) ===")
    xs_kill_ebarcode()
    for exe_name in ("QubeDrive.exe", "EBarcode.exe"):
        try:
            subprocess.run(["taskkill", "/F", "/IM", exe_name], capture_output=True, timeout=5)
            log.info(f"  OK: taskkill {exe_name} (kalau memang sedang berjalan)")
        except Exception as e:
            log.warning(f"  Gagal taskkill {exe_name}: {e}")
    time.sleep(1)
    log.info("=== PRE-START selesai ===")


def main():
    try:
        os.chdir(BASE_DIR)
    except Exception:
        pass

    log.info("=" * 60)
    log.info("  UNIFIED AUTO — XShelf -> CheckDisc -> PrintAuto24 (1 skrip)")
    log.info("=" * 60)

    progress = load_progress()
    stages = progress.get("stages", {})

    if stages.get("all_done"):
        log.info(f"[RESUME] Semua tahap untuk hari ini ({progress['date']}) SUDAH SELESAI "
                  f"(lihat {PROGRESS_FILE}). Tidak dijalankan ulang.")
        sys.exit(0)

    if stages.get("no_change"):
        log.info(f"[RESUME] CheckDisc hari ini ({progress['date']}) sudah dicek dan "
                  f"TIDAK ADA PERUBAHAN. Tidak perlu diulang.")
        sys.exit(0)

    if stages:
        log.info(f"[RESUME] Ditemukan progress hari ini ({progress['date']}): {stages}")
        log.info("[RESUME] Tahap yang sudah sukses akan DILEWATI, lanjut dari tahap yang belum/gagal.")

    close_leftover_apps_before_start()

    # TAHAP 1: XShelf
    if not stages.get("xshelf"):
        log.info("### TAHAP 1: XShelf (belum selesai - jalankan/ulangi) ###")
        ok1 = stage_xshelf()
        if not ok1:
            log.error("*** XShelf selesai dengan error. Proses dihentikan. "
                      "Jalankan skrip ini lagi kapan saja untuk melanjutkan.")
            sys.exit(1)
        mark_stage_done(progress, "xshelf")
        gap()
    else:
        log.info("### TAHAP 1: XShelf sudah selesai sebelumnya hari ini - DILEWATI ###")

    # TAHAP 2: CheckDisc
    if not stages.get("checkdisc"):
        log.info("### TAHAP 2: CheckDisc (belum selesai - jalankan/ulangi) ###")
        ok2, ada_perubahan = stage_checkdisc()
        if not ok2:
            log.error("*** CheckDisc gagal. Proses dihentikan. "
                      "Jalankan skrip ini lagi kapan saja untuk melanjutkan.")
            sys.exit(1)
        mark_stage_done(progress, "checkdisc")
        if not ada_perubahan:
            log.info("CheckDisc: tidak ada perubahan terdeteksi. Menutup otomatis.")
            progress.setdefault("stages", {})["no_change"] = True
            progress.setdefault("stages", {})["all_done"] = True
            save_progress(progress)
            sys.exit(0)
        gap()
    else:
        log.info("### TAHAP 2: CheckDisc sudah selesai sebelumnya hari ini - DILEWATI ###")

    # TAHAP 3: PrintAuto24 (trimmed)
    if not stages.get("printauto24"):
        log.info("### TAHAP 3: PrintAuto24 (belum selesai - jalankan/ulangi) ###")
        ok3 = stage_printauto24()
        if not ok3:
            log.error("*** PrintAuto24 (trimmed) selesai dengan error. "
                      "Jalankan skrip ini lagi kapan saja untuk melanjutkan.")
            sys.exit(1)
        mark_stage_done(progress, "printauto24")
        gap()
    else:
        log.info("### TAHAP 3: PrintAuto24 sudah selesai sebelumnya hari ini - DILEWATI ###")

    # TAHAP 4: Print QTY MINUS (COLLECTING LABEL) - hanya jalan kalau ada
    # file "QTY MINUS (tanggal hari ini).txt"
    if not stages.get("qty_minus"):
        log.info("### TAHAP 4: Print QTY MINUS (belum selesai - jalankan/ulangi) ###")
        ok4 = stage_print_qty_minus()
        if not ok4:
            log.error("*** Print QTY MINUS selesai dengan error. "
                      "Jalankan skrip ini lagi kapan saja untuk melanjutkan.")
            sys.exit(1)
        mark_stage_done(progress, "qty_minus")
        gap()
    else:
        log.info("### TAHAP 4: Print QTY MINUS sudah selesai sebelumnya hari ini - DILEWATI ###")

    progress.setdefault("stages", {})["all_done"] = True
    save_progress(progress)

    exit_with_countdown()
    sys.exit(0)


if __name__ == "__main__":
    main()
