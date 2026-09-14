"""
UNIFIED AUTO — Satu skrip gabungan
===================================
Urutan proses (otomatis, jeda 5 detik antar-tahap):
  TAHAP 1: XShelf         -> logikanya PERSIS SAMA seperti XShelf.py asli
                              (QubeDrive download, RUN#1 & RUN#2 export SHELF)
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
QUBEDRIVE_ARGS  = ["-Jdownload_data"]
QUBEDRIVE_WAIT  = 5 * 60     # 5 menit
TIMEOUT         = 15         # detik, timeout tunggu window
WAIT_TIMER      = 20 * 60    # 20 menit, tunggu setelah klik Search di XShelf

CHECKDISC_INPUT_PATH  = r"D:\QasDev\QubeV10\BackEnd\Data\Seuic\02\02.txt"
CHECKDISC_OUTPUT_PATH = r"D:\Scanning"
PRICE_HISTORY_FILE    = "price_history.json"
DATE_FORMAT            = "%d/%m/%Y"
EXPIRED_LOOKBACK_DAYS  = 3   # rentang "baru selesai promo" -> PRICE ADJUSTMENT (dulu 30 hari/±2 minggu)

GAP_SECONDS         = 5     # jeda antar tahap
EXIT_DELAY_SUCCESS  = 10    # tunggu sebelum exit setelah semua sukses

BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)))
LOG_FILE = os.path.join(BASE_DIR, "unified_auto_log.txt")

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


def xs_run_qubedrive_download() -> bool:
    log.info("[XShelf] STEP 0 - Jalankan QubeDrive.exe -Jdownload_data")
    try:
        subprocess.Popen([QUBEDRIVE_EXE] + QUBEDRIVE_ARGS)
        log.info(f"  OK: QubeDrive dijalankan ({QUBEDRIVE_EXE} {' '.join(QUBEDRIVE_ARGS)})")
    except Exception as e:
        log.error(f"*** FAIL: Gagal menjalankan QubeDrive.exe: {e}")
        return False

    log.info(f"[XShelf] STEP 0b - Menunggu {QUBEDRIVE_WAIT // 60} menit setelah QubeDrive dijalankan...")
    start_wait = time.time()
    while time.time() - start_wait < QUBEDRIVE_WAIT:
        elapsed_min = int((time.time() - start_wait) // 60)
        remaining_min = (QUBEDRIVE_WAIT // 60) - elapsed_min
        if elapsed_min > 0 and (time.time() - start_wait) % 60 < 5:
            log.info(f"  ... menunggu QubeDrive, sisa ~{remaining_min} menit")
        time.sleep(5)
    log.info("  OK: Selesai menunggu QubeDrive")
    return True


def xs_fail(msg: str) -> bool:
    log.error(f"*** [XShelf] FAIL: {msg}")
    xs_kill_ebarcode()
    return False


def xs_run_flow(run_qubedrive: bool = True, include_checkbox5: bool = True) -> bool:
    log.info("=== [XShelf] START EBarcode ===")

    if run_qubedrive:
        if not xs_run_qubedrive_download():
            return xs_fail("Gagal menjalankan QubeDrive.exe -Jdownload_data")

    leftover = Desktop(backend="win32").windows(class_name="TfrmBarcode")
    if leftover:
        log.warning("  [XShelf] EBarcode masih terbuka dari sesi sebelumnya, tutup dulu...")
        xs_kill_ebarcode()
        time.sleep(0.75)

    log.info("[XShelf] STEP 1 - Buka EBarcode")
    try:
        Application(backend="win32").connect(path=EXE_PATH)
        log.info("  EBarcode sudah berjalan")
    except Exception:
        log.info("  Membuka EBarcode baru...")
        subprocess.Popen([EXE_PATH])
        time.sleep(2)

    if not xs_wait_win("TfrmBarcode"):
        return xs_fail("TfrmBarcode tidak muncul")

    app = xs_get_app()
    frm = app.window(class_name="TfrmBarcode")
    xs_safe_set_focus(frm)
    time.sleep(0.35)
    log.info("  OK: TfrmBarcode aktif")

    log.info("[XShelf] STEP 2 - Klik Print (TButton3)")
    try:
        frm.child_window(class_name="TButton", found_index=2).click()
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
    try:
        app = xs_get_app()
        print_frm = app.window(class_name="TfrmPrint")
        xs_safe_set_focus(print_frm)
        time.sleep(0.2)
        print_frm.child_window(title="Export SHELF", class_name="TcxButton").click()
        time.sleep(0.5)
        log.info("  OK: Export SHELF diklik")
    except Exception as e:
        return xs_fail(f"Gagal klik Export SHELF: {e}")

    log.info("[XShelf] STEP 3 - Tunggu dialog Export | StoreSHELF")
    export_win = xs_wait_win("TdlgExportSHELF")
    if not export_win:
        return xs_fail("Dialog TdlgExportSHELF (Export | StoreSHELF) tidak muncul")
    log.info("  OK: Dialog TdlgExportSHELF muncul")

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
            chk.click()
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
        dlg.child_window(class_name="TcxButton", found_index=3).click()
        time.sleep(0.5)
        log.info("  OK: Search diklik")
    except Exception as e:
        return xs_fail(f"Gagal klik TcxButton4 (Search): {e}")

    log.info(f"[XShelf] STEP 4 - Menunggu {WAIT_TIMER // 60} menit setelah Search...")
    start_wait = time.time()
    while time.time() - start_wait < WAIT_TIMER:
        elapsed_min = int((time.time() - start_wait) // 60)
        remaining_min = (WAIT_TIMER // 60) - elapsed_min
        if elapsed_min > 0 and (time.time() - start_wait) % 60 < 5:
            log.info(f"  ... menunggu, sisa ~{remaining_min} menit")
        time.sleep(5)
    log.info("  OK: Selesai menunggu 30 menit")

    log.info("[XShelf] STEP 5 - Tunggu popup Qube Barcode Module (Done)")
    popup = xs_wait_win("TMessageForm", timeout=TIMEOUT)
    if not popup:
        return xs_fail("Popup TMessageForm (Done) tidak muncul setelah menunggu")

    try:
        app = xs_get_app()
        msg_dlg = app.window(class_name="TMessageForm")
        xs_safe_set_focus(msg_dlg)
        time.sleep(0.2)
        msg_dlg.child_window(class_name="TButton", found_index=0).click()
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
        dlg.child_window(class_name="TcxButton", found_index=2).click()
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
            "end_date": cd_parse_date(row[14].strip()),
            "col23": row[22].strip() if len(row) > 22 else "",
            "row_num": row_num,
        })
    log.info(f"[CheckDisc] Baca {len(raw_list)} baris valid (repaired={repaired_count}, "
             f"merged={merged_count}, skipped={skipped_count})")
    return raw_list


def cd_deduplicate_by_desc(raw_list):
    desc_groups = defaultdict(list)
    for item in raw_list:
        desc_groups[item["desc"]].append(item)
    kept = []
    removed = []
    for desc, items in desc_groups.items():
        if len(items) == 1:
            kept.append(items[0])
        else:
            sku_7 = [it for it in items if cd_is_7digit(it["sku"])]
            if sku_7:
                kept.append(sku_7[0])
                for it in items:
                    if it not in sku_7[:1]:
                        removed.append(it)
            else:
                for it in items:
                    kept.append(it)
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


def cd_sort_by_shelf(data):
    return sorted(data, key=lambda x: x.get("shelf", "").strip().lower())


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


def cd_export_price_adjustment(expired_items, lookback_days=EXPIRED_LOOKBACK_DAYS):
    """BEST PRICE / 'promo sudah tidak berjalan' check — turunan dari
    show_expired_alert() di CheckDisc.py asli. Mengambil item yang
    end_date-nya sudah lewat DAN selisihnya dari hari ini <= lookback_days
    (semula 30 hari / ~2 minggu, sekarang di-set 3 hari), lalu export ke
    PRICE ADJUSTMENT (tanggal).txt."""
    if not expired_items:
        return None
    today = datetime.now()
    recent_expired = [
        item for item in expired_items
        if item.get("end_date") is not None
        and item["end_date"] < today
        and (today - item["end_date"]) <= timedelta(days=lookback_days)
    ]
    if not recent_expired:
        return None
    data = cd_sort_by_shelf(recent_expired)
    lines = [item["sku"] + ",1" for item in data]
    path = cd_auto_save("PRICE ADJUSTMENT", lines)
    log.info(f"[CheckDisc] Export PRICE ADJUSTMENT (promo baru berakhir, "
             f"<= {lookback_days} hari): {len(data)} item -> {path}")
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

    dedup_data, removed = cd_deduplicate_by_desc(raw_data)
    log.info(f"[CheckDisc] Dedup: {len(dedup_data)} data, {len(removed)} dibuang")

    berjalan, expired = cd_compute_berjalan_expired(dedup_data)
    log.info(f"[CheckDisc] Aktif (berjalan): {len(berjalan)} | Sudah lewat (expired): {len(expired)}")

    price_increases, price_decreases, desc_changes = cd_check_data_changes(dedup_data)
    log.info(f"[CheckDisc] Price naik: {len(price_increases)} | Price turun: {len(price_decreases)} | "
             f"Deskripsi berubah: {len(desc_changes)}")

    exported = []

    # BEST PRICE / cek item yang promo-nya baru selesai (dalam EXPIRED_LOOKBACK_DAYS hari terakhir)
    if expired:
        p = cd_export_price_adjustment(expired)
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
PA_ALLOWED_BASENAMES = [
    "PRICE ADJUSTMENT",
    "PRICE CHANGE",
    "CHANGE DESCRIPTION",
    "NO SHELF",
]


def pa_today_allowed_items():
    date_str = datetime.now().strftime("%d-%m-%Y")
    return {f"{base} ({date_str}).TXT".upper() for base in PA_ALLOWED_BASENAMES}


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


def pa_fail(msg: str) -> bool:
    log.error(f"*** [PrintAuto24] FAIL: {msg}")
    pa_kill_ebarcode()
    return False


def stage_printauto24() -> bool:
    log.info("### TAHAP 3 - PrintAuto24 (trimmed: Step 1,2,3-Zebex,4-TdlgData) ###")

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

    if not pa_wait_win("TfrmBarcode"):
        return pa_fail("TfrmBarcode tidak muncul")

    app = pa_get_app()
    frm = app.window(class_name="TfrmBarcode")
    pa_safe_set_focus(frm)
    time.sleep(0.35)
    log.info("  OK: TfrmBarcode aktif")

    log.info("[PrintAuto24] STEP 2 - Klik Print (TButton3)")
    try:
        frm.child_window(class_name="TButton", found_index=2).click()
        time.sleep(0.5)
        log.info("  OK")
    except Exception as e:
        return pa_fail(f"Gagal klik TButton3: {e}")

    log.info("[PrintAuto24] STEP 3 - TfrmPrint -> klik Zebex")
    if not pa_wait_win("TfrmPrint", timeout=15):
        return pa_fail("TfrmPrint tidak muncul")

    app = pa_get_app()
    frm_print = app.window(class_name="TfrmPrint")
    pa_safe_set_focus(frm_print)
    time.sleep(0.25)

    try:
        zebex = None
        try:
            zebex = frm_print.child_window(title="Zebex", class_name="TcxButton")
        except Exception:
            zebex = None
        if zebex is None:
            for i, btn in enumerate(frm_print.children(class_name="TcxButton")):
                try:
                    if "ZEBEX" in btn.window_text().strip().upper():
                        zebex = btn
                        break
                except Exception:
                    pass
        if zebex is None:
            zebex = frm_print.child_window(class_name="TcxButton", found_index=8)
        zebex.click()
        time.sleep(0.5)
        log.info("  OK: Zebex diklik")
    except Exception as e:
        return pa_fail(f"Gagal klik Zebex: {e}")

    log.info("[PrintAuto24] STEP 4 - Dialog TdlgData")
    if not pa_wait_win("TdlgData", timeout=20):
        return pa_fail("TdlgData tidak muncul")

    app = pa_get_app()
    data_win = app.window(class_name="TdlgData")
    pa_safe_set_focus(data_win)
    time.sleep(0.3)

    allowed = pa_today_allowed_items()
    log.info(f"  Item yang boleh dicentang (hari ini): {sorted(allowed)}")

    try:
        data_win.child_window(class_name="TBitBtn", found_index=3).click()
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
        app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=6).click()
        time.sleep(1.25)
        log.info("  OK: Extract selesai")
    except Exception as e:
        return pa_fail(f"Gagal klik Extract: {e}")

    log.info("[PrintAuto24] STEP 4c - Print >> (TBitBtn6)")
    try:
        app = pa_get_app()
        app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=5).click()
        time.sleep(1.25)
        log.info("  OK: Print dikirim")
    except Exception as e:
        return pa_fail(f"Gagal klik Print: {e}")

    popup = pa_wait_win("TMessageForm", timeout=60)
    if popup:
        try:
            app = pa_get_app()
            dlg = app.window(class_name="TMessageForm")
            dlg.child_window(class_name="TButton", found_index=0).click()
            time.sleep(0.25)
            log.info("  OK: Popup print ditutup")
        except Exception as e:
            log.info(f"  Popup print error: {e}")
    else:
        log.info("  Tidak ada popup, lanjut")

    log.info("[PrintAuto24] STEP 4d - Close TdlgData (TBitBtn8)")
    try:
        app = pa_get_app()
        app.window(class_name="TdlgData").child_window(class_name="TBitBtn", found_index=7).click()
        time.sleep(0.25)
        log.info("  OK: TdlgData ditutup")
    except Exception as e:
        log.info(f"  Close TdlgData error: {e}")

    pa_kill_ebarcode()
    log.info("=== [PrintAuto24] SELESAI ===")
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


def main():
    try:
        os.chdir(BASE_DIR)
    except Exception:
        pass

    log.info("=" * 60)
    log.info("  UNIFIED AUTO — XShelf -> CheckDisc -> PrintAuto24 (1 skrip)")
    log.info("=" * 60)

    # TAHAP 1: XShelf
    ok1 = stage_xshelf()
    if not ok1:
        log.error("*** XShelf selesai dengan error. Proses dihentikan.")
        sys.exit(1)
    gap()

    # TAHAP 2: CheckDisc
    ok2, ada_perubahan = stage_checkdisc()
    if not ok2:
        log.error("*** CheckDisc gagal. Proses dihentikan.")
        sys.exit(1)
    if not ada_perubahan:
        log.info("CheckDisc: tidak ada perubahan terdeteksi. Menutup otomatis.")
        sys.exit(0)
    gap()

    # TAHAP 3: PrintAuto24 (trimmed)
    ok3 = stage_printauto24()
    if not ok3:
        log.error("*** PrintAuto24 (trimmed) selesai dengan error.")
        sys.exit(1)
    gap()

    exit_with_countdown()
    sys.exit(0)


if __name__ == "__main__":
    main()
