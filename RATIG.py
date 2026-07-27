import os
import sys
import json
import time
import logging
import sqlite3
import threading
import datetime
import shutil
import hashlib
import queue
import webbrowser

import serial
import serial.tools.list_ports

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import pandas as pd
import numpy as np


try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


# ==========================================
# PATHS & FOLDERS
# ==========================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR    = os.path.join(BASE_DIR, "data")
EXCEL_DIR   = os.path.join(DATA_DIR, "excel")
DB_DIR      = os.path.join(DATA_DIR, "database")
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
LOGS_DIR    = os.path.join(BASE_DIR, "logs")
CONFIG_DIR  = os.path.join(BASE_DIR, "config")
ASSETS_DIR  = os.path.join(BASE_DIR, "assets")
BACKUP_DIR  = os.path.join(BASE_DIR, "backups")

for d in [DATA_DIR, EXCEL_DIR, DB_DIR, REPORTS_DIR, LOGS_DIR, CONFIG_DIR, ASSETS_DIR, BACKUP_DIR]:
    os.makedirs(d, exist_ok=True)

LOG_FILE    = os.path.join(LOGS_DIR,   "application.log")
CONFIG_FILE = os.path.join(CONFIG_DIR, "settings.json")

# ==========================================
# LOGGING
# ==========================================
# Dual logging: file + console
_log_formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(funcName)s: %(message)s")
_file_handler  = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setFormatter(_log_formatter)
_stream_handler = logging.StreamHandler(sys.stdout)
_stream_handler.setFormatter(_log_formatter)
_stream_handler.setLevel(logging.WARNING)

logging.basicConfig(level=logging.INFO, handlers=[_file_handler, _stream_handler])
logging.info("RATIG App started. PID=%d", os.getpid())

# ==========================================
# CONFIG
# ==========================================
DEFAULT_CONFIG = {
    "baudrate": 115200,
    "auto_connect": True,
    "excel_file": "data/excel/ratig_data.xlsx",
    "database_file": "data/database/ratig.db",
    "company_name": "PT. NAMA PERUSAHAAN",
    "created_by": "Marsellinus, Khilqi, Idris",
    "auto_backup": True,
    "age_warning_threshold": 60,
    "daily_test_limit": 3,
    "admin_pin": "1234",
    "admin_pin_hashed": False,   # flag: apakah PIN sudah di-hash
    "paramedis_pin": "paramedis123",
    "paramedis_pin_hashed": False,  # flag: apakah PIN Paramedis sudah di-hash
    "chart_limit": 500,          # max rows untuk chart
    "table_limit": 500,          # max rows untuk tabel
    "version": "2.0",
    "app_name": "FAST"
}

import base64

_SECRET_KEY = b"ITPSHE#2026_RATIG_SECURE_V2"

def _encrypt_data(plain_str: str) -> str:
    b_data = plain_str.encode('utf-8')
    xored = bytes(a ^ _SECRET_KEY[i % len(_SECRET_KEY)] for i, a in enumerate(b_data))
    return base64.b64encode(xored).decode('utf-8')

def _decrypt_data(b64_str: str) -> str:
    try:
        b_data = base64.b64decode(b64_str)
        xored = bytes(a ^ _SECRET_KEY[i % len(_SECRET_KEY)] for i, a in enumerate(b_data))
        return xored.decode('utf-8')
    except Exception:
        return b64_str

def _save_config(config_dict: dict):
    try:
        json_str = json.dumps(config_dict, indent=4)
        encrypted = _encrypt_data(json_str)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(encrypted)
    except Exception as e:
        logging.error(f"Gagal save config: {e}")

def _load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        _save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        decrypted = _decrypt_data(content)
        return json.loads(decrypted)
    except Exception as e:
        logging.error(f"Config corrupted: {e}")
        _save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

def _hash_pin(pin: str) -> str:
    """SHA-256 hash untuk PIN admin."""
    return hashlib.sha256(pin.encode()).hexdigest()

def _verify_pin(input_pin: str) -> bool:
    """Verifikasi PIN: support plain dan hashed."""
    stored = str(CONFIG.get("admin_pin", "1234"))
    if CONFIG.get("admin_pin_hashed", False):
        return _hash_pin(input_pin) == stored
    return input_pin == stored

def _migrate_pin_to_hash():
    """Upgrade plain PIN → hashed jika belum."""
    if not CONFIG.get("admin_pin_hashed", False):
        plain = CONFIG.get("admin_pin", "1234")
        CONFIG["admin_pin"]        = _hash_pin(plain)
        CONFIG["admin_pin_hashed"] = True
        _save_config(CONFIG)

def _verify_paramedis_pin(input_pin: str) -> bool:
    """Verifikasi PIN Mode Paramedis: support plain dan hashed."""
    stored = str(CONFIG.get("paramedis_pin", "paramedis123"))
    if not stored:
        return True  # PIN dikosongkan = proteksi dimatikan
    if CONFIG.get("paramedis_pin_hashed", False):
        return _hash_pin(input_pin or "") == stored
    return input_pin == stored

def _migrate_paramedis_pin_to_hash():
    """Upgrade PIN Paramedis plain → hashed jika belum."""
    if not CONFIG.get("paramedis_pin_hashed", False):
        plain = CONFIG.get("paramedis_pin", "paramedis123")
        CONFIG["paramedis_pin"]        = _hash_pin(plain) if plain else ""
        CONFIG["paramedis_pin_hashed"] = True
        _save_config(CONFIG)
CONFIG = _load_config()
for k, v in DEFAULT_CONFIG.items():
    CONFIG.setdefault(k, v)

# LOCKED_MODE: set to "SHE Officer" atau "Paramedis" untuk membuat versi
# aplikasi yang dikunci pada satu mode (dropdown mode disembunyikan/nonaktif).
# Biarkan None untuk versi gabungan (dropdown mode aktif, bisa dipilih user).
LOCKED_MODE = None

DB_PATH       = os.path.join(BASE_DIR, os.path.normpath(CONFIG["database_file"]))
EXCEL_PATH    = os.path.join(BASE_DIR, os.path.normpath(CONFIG["excel_file"]))
EXCEL_MK_PATH = os.path.join(EXCEL_DIR, "Master_Karyawan.xlsx")

# ==========================================
# THEME
# ==========================================
CLR = {
    "bg_dark":    "#F0F2F5",
    "bg_panel":   "#FFFFFF",
    "bg_card":    "#F8F9FA",
    "bg_sidebar": "#2C3E50",
    "sidebar_txt":"#BDC3C7",
    "sidebar_sel":"#3498DB",
    "accent":     "#3498DB",
    "accent2":    "#2ECC71",
    "danger":     "#E74C3C",
    "warning":    "#F39C12",
    "success":    "#27AE60",
    "text_main":  "#2C3E50",
    "text_muted": "#7F8C8D",
    "separator":  "#DEE2E6",
    "row_even":   "#F8F9FA",
    "row_odd":    "#FFFFFF",
    "progress_bg":"#E8F4FD",
    "fatigue_colors": {
        "Normal":         "#FFFFFF",
        "Fatigue Ringan": "#2ECC71",
        "Fatigue Sedang": "#F1C40F",
        "Fatigue Berat":  "#E74C3C",
        "Error/Invalid":  "#95A5A6",
    },
    "fatigue_fg": {
        "Normal":         "#000000",
        "Fatigue Ringan": "#FFFFFF",
        "Fatigue Sedang": "#000000",
        "Fatigue Berat":  "#FFFFFF",
        "Error/Invalid":  "#FFFFFF",
    }
}

# ==========================================
# DATABASE
# ==========================================
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.executescript("""
        CREATE TABLE IF NOT EXISTS master_karyawan (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            nama           TEXT,
            nik            TEXT UNIQUE,
            usia           TEXT,
            jabatan        TEXT,
            info_pekerjaan TEXT
        );
        CREATE TABLE IF NOT EXISTS master_kontraktor (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            nama           TEXT,
            nik            TEXT UNIQUE,
            usia           TEXT,
            jabatan        TEXT,
            info_pekerjaan TEXT
        );
        CREATE TABLE IF NOT EXISTS results (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal       TEXT,
            nama          TEXT,
            nik           TEXT,
            usia          TEXT,
            jenis_kelamin TEXT,
            jabatan       TEXT,
            contractor    TEXT,
            plan          TEXT,
            divisi        TEXT,
            mode          TEXT,
            trial         TEXT,
            reaction_time REAL,
            t1 REAL, t2 REAL, t3 REAL,
            t4 REAL, t5 REAL, t6 REAL,
            avg_reaction  REAL,
            status        TEXT,
            fatigue_level TEXT,
            td_sistol     REAL,
            td_diastol    REAL,
            nadi          REAL,
            alcohol_test  REAL,
            kesimpulan_sistol  TEXT,
            kesimpulan_diastol TEXT,
            kesimpulan_nadi    TEXT,
            keputusan     TEXT,
            rekomendasi   TEXT,
            keterangan    TEXT
        );
        """)

        # Performance indexes
        c.executescript("""
        CREATE INDEX IF NOT EXISTS idx_results_nik     ON results(nik);
        CREATE INDEX IF NOT EXISTS idx_results_tanggal ON results(DATE(tanggal));
        CREATE INDEX IF NOT EXISTS idx_results_fatigue ON results(fatigue_level);
        CREATE INDEX IF NOT EXISTS idx_mk_nik          ON master_karyawan(nik);
        CREATE INDEX IF NOT EXISTS idx_mkon_nik        ON master_kontraktor(nik);
        """)

        existing = {row[1] for row in c.execute("PRAGMA table_info(results)")}
        mig = [
            ("nik",                 "TEXT"),
            ("usia",                "TEXT"),
            ("jabatan",             "TEXT"),
            ("contractor",          "TEXT"),
            ("plan",                "TEXT"),
            ("divisi",              "TEXT"),
            ("t1",                  "REAL"),
            ("t2",                  "REAL"),
            ("t3",                  "REAL"),
            ("t4",                  "REAL"),
            ("t5",                  "REAL"),
            ("t6",                  "REAL"),
            ("avg_reaction",        "REAL"),
            ("td_sistol",           "REAL"),
            ("td_diastol",          "REAL"),
            ("nadi",                "REAL"),
            ("alcohol_test",        "REAL"),
            ("kesimpulan_sistol",   "TEXT"),
            ("kesimpulan_diastol",  "TEXT"),
            ("kesimpulan_nadi",     "TEXT"),
            ("keputusan",           "TEXT"),
            ("rekomendasi",         "TEXT"),
            ("keterangan",          "TEXT"),
            ("kesimpulan_alkohol",  "TEXT"),
            ("jenis_kelamin",       "TEXT"),
            ("keluhan",             "TEXT"),
            ("diagnosis_awal",      "TEXT"),
            ("tindakan_medis",      "TEXT"),
            ("pemberian_obat",      "TEXT"),
        ]
        for col, dtype in mig:
            if col not in existing:
                try:
                    c.execute(f"ALTER TABLE results ADD COLUMN {col} {dtype}")
                except Exception:
                    pass

        mk_existing = {row[1] for row in c.execute("PRAGMA table_info(master_karyawan)")}
        for col in ("jabatan", "jenis_kelamin", "tanggal_lahir"):
            if col not in mk_existing:
                try:
                    c.execute(f"ALTER TABLE master_karyawan ADD COLUMN {col} TEXT")
                except Exception:
                    pass

        mkon_existing = {row[1] for row in c.execute("PRAGMA table_info(master_kontraktor)")}
        for col in ("jenis_kelamin", "tanggal_lahir"):
            if col not in mkon_existing:
                try:
                    c.execute(f"ALTER TABLE master_kontraktor ADD COLUMN {col} TEXT")
                except Exception:
                    pass

        conn.commit()


def init_excel():
    if not os.path.exists(EXCEL_PATH):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "No", "Tanggal", "Nama", "NIK", "Usia", "Jenis Kelamin", "Bagian",
            "T1", "T2", "T3", "T4", "T5", "T6",
            "Tekanan Darah (Sistol)", "Tekanan Darah (Diastol)",
            "Nadi", "Nilai Total HFT", "Nilai Rata-Rata HFT",
            "Hasil Alcohol Test",
            "Kesimpulan (Sistol)", "Kesimpulan (Diastol)",
            "Kesimpulan (Nadi)", "Kesimpulan (Fatigue)",
            "Keputusan", "Rekomendasi", "Keterangan", "Kesimpulan (Alkohol)"
        ])
        wb.save(EXCEL_PATH)

def update_excel_medis(db_id, data):
    def task():
        if not os.path.exists(EXCEL_PATH):
            return
        try:
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
            db_id_str = str(db_id)

            col_map = {}
            for col in range(1, ws.max_column + 1):
                header = str(ws.cell(row=1, column=col).value or "").strip()
                if header:
                    col_map[header] = col

            field_to_headers = {
                "td_sistol":          ["Tekanan Darah (Sistol)", "TD Sistol"],
                "td_diastol":         ["Tekanan Darah (Diastol)", "TD Diastol"],
                "nadi":               ["Nadi"],
                "alcohol_test":       ["Hasil Alcohol Test", "Kadar Alkohol"],
                "kesimpulan_sistol":  ["Kesimpulan (Sistol)"],
                "kesimpulan_diastol": ["Kesimpulan (Diastol)"],
                "kesimpulan_nadi":    ["Kesimpulan (Nadi)"],
                "keputusan":          ["Keputusan"],
                "rekomendasi":        ["Rekomendasi"],
                "keterangan":         ["Keterangan"],
                "kesimpulan_alkohol": ["Kesimpulan (Alkohol)"]
            }

            for row in range(2, ws.max_row + 1):
                if str(ws.cell(row=row, column=1).value) == db_id_str:
                    for field, possible_headers in field_to_headers.items():
                        val = data.get(field, "")
                        for h in possible_headers:
                            if h in col_map:
                                ws.cell(row=row, column=col_map[h]).value = val
                                break
                    wb.save(EXCEL_PATH)
                    logging.debug(f"Excel medis updated for id={db_id}")
                    break
        except Exception as e:
            logging.error(f"Excel medis update: {e}")
    threading.Thread(target=task, daemon=True).start()

# ==========================================
# AUTO BACKUP
# ==========================================
def auto_backup_if_needed():
    if not CONFIG.get("auto_backup", True):
        return
    today  = datetime.date.today().strftime("%Y-%m-%d")
    marker = os.path.join(BACKUP_DIR, f".backup_{today}")
    if os.path.exists(marker):
        return
    try:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if os.path.exists(DB_PATH):
            shutil.copy2(DB_PATH,    os.path.join(BACKUP_DIR, f"ratig_{ts}.db"))
        if os.path.exists(EXCEL_PATH):
            shutil.copy2(EXCEL_PATH, os.path.join(BACKUP_DIR, f"ratig_{ts}.xlsx"))
        with open(marker, "w") as f:
            f.write(ts)
        logging.info(f"Auto-backup done: {ts}")
    except Exception as e:
        logging.error(f"Auto-backup failed: {e}")


def auto_backup_loop():
    last_mtime = 0
    realtime_db_path = os.path.join(BACKUP_DIR, "ratig_realtime.db")
    realtime_excel_path = os.path.join(BACKUP_DIR, "ratig_data_realtime.xlsx")
    
    while True:
        try:
            # 1. Snapshot Harian (Daily Backup)
            auto_backup_if_needed()
            
            # 2. Realtime Mirroring (Sangat Efisien)
            if CONFIG.get("auto_backup", True) and os.path.exists(DB_PATH):
                current_mtime = os.path.getmtime(DB_PATH)
                if current_mtime > last_mtime:
                    try:
                        # Menggunakan API resmi SQLite untuk backup page-by-page tanpa mengunci database
                        with sqlite3.connect(DB_PATH) as src, sqlite3.connect(realtime_db_path) as dst:
                            src.backup(dst)
                        
                        # Full Export Excel Backup (agar "ban serep" memiliki semua data dari awal)
                        try:
                            conn = sqlite3.connect(DB_PATH)
                            df = pd.read_sql_query(
                                "SELECT id, tanggal, nama, nik, usia, jenis_kelamin, jabatan, "
                                "IFNULL(t1,0) AS t1, IFNULL(t2,0) AS t2, IFNULL(t3,0) AS t3, "
                                "IFNULL(t4,0) AS t4, IFNULL(t5,0) AS t5, IFNULL(t6,0) AS t6, "
                                "IFNULL(td_sistol,'') AS sistol, IFNULL(td_diastol,'') AS diastol, "
                                "IFNULL(nadi,'') AS nadi, "
                                "(IFNULL(t1,0)+IFNULL(t2,0)+IFNULL(t3,0)+IFNULL(t4,0)+IFNULL(t5,0)+IFNULL(t6,0)) AS total, "
                                "avg_reaction, IFNULL(alcohol_test,'') AS alcohol_test, "
                                "IFNULL(kesimpulan_sistol,'') AS kesimpulan_sistol, "
                                "IFNULL(kesimpulan_diastol,'') AS kesimpulan_diastol, "
                                "IFNULL(kesimpulan_nadi,'') AS kesimpulan_nadi, "
                                "IFNULL(fatigue_level,'') AS fatigue_level, "
                                "IFNULL(keputusan,'') AS keputusan, IFNULL(rekomendasi,'') AS rekomendasi, "
                                "IFNULL(keterangan,'') AS keterangan, IFNULL(kesimpulan_alkohol,'') AS kesimpulan_alkohol "
                                "FROM results ORDER BY id ASC", conn)
                            conn.close()
                            
                            try:
                                df["usia"] = pd.to_numeric(df["usia"], errors='coerce').astype('Int64')
                            except: pass
                            
                            col_map = {
                                "id": "No", "tanggal": "Tanggal", "nama": "Nama", "nik": "NIK", 
                                "usia": "Usia", "jenis_kelamin": "Jenis Kelamin", "jabatan": "Bagian",
                                "t1": "T1", "t2": "T2", "t3": "T3", "t4": "T4", "t5": "T5", "t6": "T6",
                                "sistol": "Tekanan Darah (Sistol)", "diastol": "Tekanan Darah (Diastol)", "nadi": "Nadi",
                                "total": "Nilai Total HFT", "avg_reaction": "Nilai Rata-Rata HFT",
                                "alcohol_test": "Hasil Alcohol Test",
                                "kesimpulan_sistol": "Kesimpulan (Sistol)",
                                "kesimpulan_diastol": "Kesimpulan (Diastol)",
                                "kesimpulan_nadi": "Kesimpulan (Nadi)",
                                "fatigue_level": "Kesimpulan (Fatigue)",
                                "keputusan": "Keputusan", "rekomendasi": "Rekomendasi",
                                "keterangan": "Keterangan", "kesimpulan_alkohol": "Kesimpulan (Alkohol)"
                            }
                            df.rename(columns=col_map, inplace=True)
                            df.to_excel(realtime_excel_path, index=False)
                        except Exception as excel_err:
                            logging.error(f"Realtime excel export error: {excel_err}")
                            
                        last_mtime = current_mtime
                    except Exception as e:
                        logging.error(f"Realtime backup error: {e}")
        except Exception as e:
            logging.error(f"Auto-backup loop error: {e}")
        
        time.sleep(3)  # Cek perubahan setiap 3 detik


# ==========================================
# HELPER: fatigue level dari avg reaction time
# Batas: Normal <240ms, Ringan 240–400ms, Sedang 401–580ms, Berat ≥581ms
# ==========================================
def classify_fatigue(avg: float) -> str:
    """Klasifikasi level fatigue berdasarkan rata-rata reaction time (ms)."""
    if avg <= 0:   return "Error/Invalid"
    if avg < 240:  return "Normal"
    if avg <= 400: return "Fatigue Ringan"
    if avg <= 580: return "Fatigue Sedang"
    return "Fatigue Berat"


def _open_file(path: str):
    """Buka file dengan aplikasi default OS."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception as e:
        logging.warning(f"Gagal membuka file: {e}")


# ==========================================
# WIDGETS
# ==========================================
class HoverButton(tk.Button):
    def __init__(self, master, **kw):
        super().__init__(master=master, **kw)
        self._bg = self.cget("background")
        self.bind("<Enter>", lambda _: self.configure(bg=self._lighten(self._bg)))
        self.bind("<Leave>", lambda _: self.configure(bg=self._bg))

    @staticmethod
    def _lighten(h, a=0.15):
        try:
            r, g, b = int(h[1:3], 16), int(h[3:5], 16), int(h[5:7], 16)
            return (f"#{int(r+(255-r)*a):02x}"
                    f"{int(g+(255-g)*a):02x}"
                    f"{int(b+(255-b)*a):02x}")
        except Exception:
            return h


class Tooltip:
    """Tooltip sederhana — muncul saat hover widget."""
    def __init__(self, widget, text: str, delay: int = 700):
        self.widget = widget
        self.text   = text
        self.delay  = delay
        self._tw    = None
        self._job   = None
        widget.bind("<Enter>",    self._schedule)
        widget.bind("<Leave>",    self._cancel)
        widget.bind("<Button>",   self._cancel)

    def _schedule(self, _=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self, _=None):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None
        if self._tw:
            self._tw.destroy()
            self._tw = None

    def _show(self):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tw = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left",
                 background="#FFFFE0", relief="solid", borderwidth=1,
                 font=("Segoe UI", 8), padx=6, pady=3
                 ).pack()





class HistoryPopup(tk.Toplevel):
    """Popup riwayat tes per-NIK, dengan grafik trend."""

    def __init__(self, parent, nik, nama):
        super().__init__(parent)
        self.title(f"Riwayat Tes — {nama} ({nik})")
        self.geometry("820x560")
        self.configure(bg=CLR["bg_panel"])
        self.resizable(True, True)
        self.grab_set()

        try:
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query(
                "SELECT tanggal, avg_reaction, fatigue_level, divisi "
                "FROM results WHERE nik=? ORDER BY tanggal ASC",
                conn, params=(nik,))
            conn.close()
        except Exception as e:
            tk.Label(self, text=f"Error: {e}", bg=CLR["bg_panel"]).pack()
            return

        # Header
        hdr = tk.Frame(self, bg=CLR["accent"], padx=14, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text=f"📊 {nama}  |  NIK: {nik}  |  Total tes: {len(df)}",
                 font=("Segoe UI", 11, "bold"),
                 fg="white", bg=CLR["accent"]).pack(side="left")

        # Check 3x consecutive Fatigue Berat
        consecutive_berat = 0
        for level in df["fatigue_level"].values:
            if level == "Fatigue Berat":
                consecutive_berat += 1
            else:
                consecutive_berat = 0

        if consecutive_berat >= 3:
            tk.Label(
                self,
                text="⚠ PERINGATAN: Karyawan ini mengalami Fatigue Berat "
                     "3x berturut-turut pada tes terakhir!",
                font=("Segoe UI", 10, "bold"),
                fg=CLR["danger"], bg=CLR["bg_panel"]
            ).pack(pady=(8, 0))

        if df.empty:
            tk.Label(self,
                     text="Belum ada data tes untuk karyawan ini.",
                     font=("Segoe UI", 11), fg=CLR["text_muted"],
                     bg=CLR["bg_panel"]).pack(expand=True)
            return

        # Chart
        fig = Figure(figsize=(7.5, 2.8), dpi=96, facecolor=CLR["bg_panel"])
        ax  = fig.add_subplot(111)
        ax.set_facecolor(CLR["bg_card"])
        ax.plot(range(len(df)), df["avg_reaction"].values,
                color=CLR["accent"], marker="o", markersize=5, linewidth=1.8)
        ax.axhline(240, color=CLR["success"], linestyle="--",
                   linewidth=1, alpha=.7, label="Normal <240")
        ax.axhline(401, color=CLR["warning"], linestyle="--",
                   linewidth=1, alpha=.7, label="Sedang <401")
        ax.axhline(581, color=CLR["danger"],  linestyle="--",
                   linewidth=1, alpha=.7, label="Berat <581")
        ax.set_title("Trend Reaction Time (ms)", fontsize=9,
                     color=CLR["text_main"])
        ax.set_ylabel("ms", fontsize=8, color=CLR["text_muted"])
        ax.tick_params(colors=CLR["text_muted"], labelsize=7)
        for sp in ax.spines.values():
            sp.set_color(CLR["separator"])
        ax.legend(fontsize=7, facecolor=CLR["bg_card"],
                  edgecolor=CLR["separator"], labelcolor=CLR["text_muted"])
        step = max(1, len(df) // 8)
        ax.set_xticks(range(0, len(df), step))
        ax.set_xticklabels(
            [str(df["tanggal"].iloc[i])[:10] for i in range(0, len(df), step)],
            rotation=25, fontsize=6)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=self).get_tk_widget().pack(
            fill="x", padx=10, pady=(8, 0))

        # Table
        frm  = tk.Frame(self, bg=CLR["bg_panel"])
        frm.pack(fill="both", expand=True, padx=10, pady=6)
        cols = ("Tanggal", "Avg RT (ms)", "Fatigue", "Divisi")
        tv   = ttk.Treeview(frm, columns=cols, show="headings", height=8)
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, anchor="center", width=180)
        for i, (_, row) in enumerate(df.iterrows()):
            fl = row["fatigue_level"] or "-"
            tv.insert("", "end",
                      values=(str(row["tanggal"])[:19],
                              f"{row['avg_reaction']:.1f}"
                              if row["avg_reaction"] else "-",
                              fl,
                              row["divisi"] or "-"),
                      tags=("e" if i % 2 == 0 else "o",))
        tv.tag_configure("e", background=CLR["row_even"])
        tv.tag_configure("o", background=CLR["row_odd"])
        sb = ttk.Scrollbar(frm, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")


class ManualInputDialog(tk.Toplevel):
    """Input data tes secara manual tanpa hardware."""

    def __init__(self, parent, on_save):
        super().__init__(parent)
        self.title("📝 Input Manual")
        self.geometry("400x500")
        self.configure(bg=CLR["bg_panel"])
        self.resizable(False, False)
        self.grab_set()
        self.on_save = on_save

        tk.Label(self, text="Input Data Tes Manual",
                 font=("Segoe UI", 13, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(14, 8))

        form = tk.Frame(self, bg=CLR["bg_panel"])
        form.pack(padx=20, fill="x")

        self._vars = {}
        fields = [
            ("NIK",               "nik",     False),
            ("Nama",              "nama",    False),
            ("Usia",              "usia",    True),
            ("Jenis Kelamin",     "jenis_kelamin", False),
            ("Jabatan",           "jabatan", False),
            ("Perusahaan/Bagian", "divisi",  False),
            ("T1 (ms)",           "t1",      True),
            ("T2 (ms)",           "t2",      True),
            ("T3 (ms)",           "t3",      True),
            ("T4 (ms)",           "t4",      True),
            ("T5 (ms)",           "t5",      True),
            ("T6 (ms)",           "t6",      True),
        ]
        vcmd = (self.register(
            lambda v: v == "" or v.replace(".", "", 1).isdigit()), "%P")

        for i, (label, key, num) in enumerate(fields):
            tk.Label(form, text=label, font=("Segoe UI", 11, "bold"),
                     fg=CLR["text_main"], bg=CLR["bg_panel"],
                     width=14, anchor="w").grid(row=i, column=0, pady=3)
            v = tk.StringVar()
            e = tk.Entry(form, textvariable=v, font=("Segoe UI", 11),
                         bg="#F1F5F9", fg=CLR["text_main"],
                         relief="flat", bd=0, width=24,
                         highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                         highlightthickness=1)
            if num:
                e.configure(validate="key", validatecommand=vcmd)
            e.grid(row=i, column=1, pady=3, padx=4)
            self._vars[key] = v

        btn_f = tk.Frame(self, bg=CLR["bg_panel"])
        btn_f.pack(pady=12)
        HoverButton(btn_f, text="Simpan", command=self._save,
                    font=("Segoe UI", 10, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", padx=18, pady=5,
                    cursor="hand2").pack(side="left", padx=6)
        HoverButton(btn_f, text="Batal", command=self.destroy,
                    font=("Segoe UI", 10),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", padx=18, pady=5,
                    cursor="hand2").pack(side="left", padx=6)

    def _save(self):
        try:
            nik = self._vars["nik"].get().strip().upper()
            if not nik:
                messagebox.showwarning("Peringatan",
                                       "NIK tidak boleh kosong!", parent=self)
                return
            if len(nik) < 3:
                messagebox.showwarning("Peringatan",
                                       "NIK terlalu pendek (minimal 3 karakter)!", parent=self)
                return

            nama = self._vars["nama"].get().strip()
            if not nama:
                messagebox.showwarning("Peringatan",
                                       "Nama tidak boleh kosong!", parent=self)
                return

            try:
                usia_val = int(self._vars["usia"].get())
                if not (1 <= usia_val <= 120):
                    raise ValueError()
                thr = CONFIG.get("age_warning_threshold", 60)
                if usia_val >= thr:
                    if not messagebox.askyesno(
                        "Peringatan Usia",
                        f"Karyawan ini berusia {usia_val} tahun (Batas: {thr} tahun).\n\n"
                        "Tetap lanjutkan menyimpan data?", parent=self):
                        return
            except ValueError:
                messagebox.showwarning("Peringatan", "Usia harus angka antara 1–120.", parent=self)
                return

            today = datetime.datetime.now().strftime("%Y-%m-%d")
            limit = CONFIG.get("daily_test_limit", 3)
            with sqlite3.connect(DB_PATH) as conn:
                tests_today = conn.execute(
                    "SELECT COUNT(*) FROM results WHERE nik=? AND tanggal LIKE ?",
                    (nik, today + "%")).fetchone()[0]
            if tests_today >= limit:
                messagebox.showwarning(
                    "Tes Ditolak",
                    f"NIK {nik} sudah dites {tests_today}x hari ini!\n"
                    f"Berdasarkan kebijakan, satu orang hanya boleh tes maksimal {limit}x per hari.", parent=self)
                return

            ts   = [float(self._vars[f"t{i}"].get() or 0) for i in range(1, 7)]
            vals = [t for t in ts if 0 < t <= 2000]
            if len(vals) < 3:
                messagebox.showwarning("Peringatan",
                    "Minimal 3 nilai T yang valid (0 < T ≤ 2000 ms) diperlukan.", parent=self)
                return
            avg  = round(sum(vals) / len(vals), 1)
            fl   = classify_fatigue(avg)

            row = {
                "tanggal":      datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "nama":         nama,
                "nik":          nik,
                "usia":         self._vars["usia"].get(),
                "jenis_kelamin": self._vars["jenis_kelamin"].get(),
                "jabatan":      self._vars["jabatan"].get(),
                "divisi":       self._vars["divisi"].get(),
                "ts":           ts,
                "avg":          avg,
                "fatigue":      fl,
            }
            self.on_save(row)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)


class MedisInputDialog(tk.Toplevel):
    """Input data medis untuk riwayat tes yang sudah ada."""

    def __init__(self, parent, db_id, on_save, app_mode="Paramedis"):
        super().__init__(parent)
        self.title("🩺 Input Medis")
        # Fit vertically within standard 768p screens (650px height is ideal, using scrollbar for overflow)
        self.geometry("450x650" if app_mode == "Paramedis" else "450x500")
        self.configure(bg=CLR["bg_panel"])
        self.resizable(False, True)
        self.grab_set()
        self.db_id   = db_id
        self.on_save = on_save
        self.app_mode = app_mode

        tk.Label(self, text="Input Data Medis & Rekomendasi",
                 font=("Segoe UI", 13, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(14, 8))

        canvas = tk.Canvas(self, bg=CLR["bg_panel"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        form = tk.Frame(canvas, bg=CLR["bg_panel"])
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True, padx=20)
        canvas.create_window((0, 0), window=form, anchor="nw")
        
        def _on_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        form.bind("<Configure>", _on_configure)

        self._vars = {}
        
        self.var_rt = tk.StringVar(value="-")
        self.var_ft = tk.StringVar(value="-")
        
        tk.Label(form, text="Rata-rata RT (ms)", font=("Segoe UI", 11, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"], width=15, anchor="w").grid(row=0, column=0, pady=3)
        tk.Entry(form, textvariable=self.var_rt, font=("Segoe UI", 11, "bold"),
                 bg=CLR["bg_panel"], fg=CLR["accent"], relief="flat", state="readonly", width=22).grid(row=0, column=1, pady=3, padx=4)

        tk.Label(form, text="Level Fatigue", font=("Segoe UI", 11, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"], width=15, anchor="w").grid(row=1, column=0, pady=3)
        tk.Entry(form, textvariable=self.var_ft, font=("Segoe UI", 11, "bold"),
                 bg=CLR["bg_panel"], fg=CLR["accent"], relief="flat", state="readonly", width=22).grid(row=1, column=1, pady=3, padx=4)

        num_fields = []
        if self.app_mode == "Paramedis":
            num_fields.extend([
                ("TD Sistol",    "td_sistol"),
                ("TD Diastol",   "td_diastol"),
                ("Nadi",         "nadi"),
            ])
        num_fields.append(("Alcohol Test (BAC%)", "alcohol_test"))

        vcmd = (self.register(
            lambda v: v == "" or v.replace(".", "", 1).isdigit()), "%P")

        for i, (label, key) in enumerate(num_fields):
            f_lbl = ("Segoe UI", 11, "bold")
            f_ent = ("Segoe UI", 11)
            fg_lbl = CLR["text_main"]
            w_lbl = 15

            tk.Label(form, text=label, font=f_lbl,
                     fg=fg_lbl, bg=CLR["bg_panel"],
                     width=w_lbl, anchor="w").grid(row=i+2, column=0, pady=3)
            v = tk.StringVar()
            e = tk.Entry(form, textvariable=v, font=f_ent,
                         bg="#F1F5F9", fg=CLR["text_main"],
                         relief="flat", bd=0, width=22,
                         highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                         highlightthickness=1)
            e.configure(validate="key", validatecommand=vcmd)
            
            # Auto-calculate triggers for instant updates
            e.bind("<KeyRelease>", lambda event: self._auto_calc())
            e.bind("<Return>", lambda event: self._auto_calc())

            if key == "alcohol_test":
                def fmt_alc(event, var=v):
                    val = var.get().strip()
                    if val:
                        try: var.set(f"{float(val):.2f}")
                        except: pass
                    self._auto_calc()
                e.bind("<FocusOut>", fmt_alc)
            else:
                e.bind("<FocusOut>", lambda event: self._auto_calc())

            e.grid(row=i+2, column=1, pady=3, padx=4)
            self._vars[key] = v

        sep_row = len(num_fields) + 2
        tk.Frame(form, bg=CLR["separator"], height=1).grid(
            row=sep_row, column=0, columnspan=2, pady=10, sticky="ew")

        start_idx = sep_row + 1
        dropdowns = []
        if self.app_mode == "Paramedis":
            dropdowns.extend([
                ("Kesimpulan Sistol",  "kesimpulan_sistol",
                 ["Hipotensi", "Normal", "Hypertensi Ringan",
                  "Hypertensi Sedang", "Hypertensi Berat"]),
                ("Kesimpulan Diastol", "kesimpulan_diastol",
                 ["Hipotensi", "Normal", "Hypertensi Ringan",
                  "Hypertensi Sedang", "Hypertensi Berat"]),
                ("Kesimpulan Nadi",    "kesimpulan_nadi",
                 ["Bradikardia", "Normal", "Takikardia Ringan",
                  "Takikardia Sedang", "Takikardia Berat"]),
            ])

        dropdowns.append(
            ("Kesimpulan Alkohol", "kesimpulan_alkohol",
             ["Lolos", "Tidak Lolos"]))

        # "Keputusan" & "Rekomendasi" = kesimpulan pemeriksaan yang menjadi
        # dasar apakah pekerja boleh bekerja -> khusus Mode Paramedis (sesuai spesifikasi).
        if self.app_mode == "Paramedis":
            dropdowns.extend([
                ("Keputusan",          "keputusan",
                 ["Diizinkan bekerja",
                  "Diizinkan bekerja, dengan catatan",
                  "Tidak diizinkan bekerja",
                  "Rujuk ke Faskes"]),
                ("Rekomendasi",        "rekomendasi",
                 ["Manajemen Fatigue", "Manajemen Hipertensi",
                  "Manajemen Fatigue & Hipertensi",
                  "Istirahat", "Lainnya"]),
            ])

        for i, (label, key, vals) in enumerate(dropdowns):
            r = start_idx + i
            tk.Label(form, text=label, font=("Segoe UI", 11, "bold"),
                     fg=CLR["text_main"], bg=CLR["bg_panel"],
                     width=15, anchor="w").grid(row=r, column=0, pady=3)
            v = tk.StringVar()
            ttk.Combobox(form, textvariable=v, values=vals,
                         font=("Segoe UI", 11), width=20,
                         state="normal").grid(row=r, column=1, pady=3, padx=4)
            self._vars[key] = v

        text_fields = []
        if self.app_mode == "Paramedis":
            pass
            
        text_row = start_idx + len(dropdowns)
        self._text_vars = {}
        for i, (label, key) in enumerate(text_fields):
            r = text_row + i
            tk.Label(form, text=label, font=("Segoe UI", 11, "bold"),
                     fg=CLR["text_main"], bg=CLR["bg_panel"],
                     width=15, anchor="nw").grid(row=r, column=0, pady=5)
            t = tk.Text(form, font=("Segoe UI", 11),
                        bg="#F1F5F9", fg=CLR["text_main"],
                        relief="flat", bd=0, width=22, height=2,
                        highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                        highlightthickness=1)
            t.grid(row=r, column=1, pady=5, padx=4)
            self._text_vars[key] = t

        ket_row = text_row + len(text_fields)
        tk.Label(form, text="Keterangan", font=("Segoe UI", 11, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"],
                 width=15, anchor="nw").grid(row=ket_row, column=0, pady=5)
        self.txt_ket = tk.Text(form, font=("Segoe UI", 11),
                               bg="#F1F5F9", fg=CLR["text_main"],
                               relief="flat", bd=0, width=22, height=3,
                               highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                               highlightthickness=1)
        self.txt_ket.grid(row=ket_row, column=1, pady=5, padx=4)

        # Load existing data before showing buttons
        self._load_data()

        btn_f = tk.Frame(form, bg=CLR["bg_panel"])
        btn_f.grid(row=ket_row+1, column=0, columnspan=2, pady=15)
        HoverButton(btn_f, text="Auto Kesimpulan", command=self._auto_calc,
                    font=("Segoe UI", 11, "bold"),
                    bg=CLR["warning"], fg="white",
                    relief="flat", padx=12, pady=6,
                    cursor="hand2").pack(side="left", padx=6)
        HoverButton(btn_f, text="Simpan", command=self._save,
                    font=("Segoe UI", 10, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", padx=18, pady=5,
                    cursor="hand2").pack(side="left", padx=4)
        HoverButton(btn_f, text="Batal", command=self.destroy,
                    font=("Segoe UI", 10),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", padx=18, pady=5,
                    cursor="hand2").pack(side="left", padx=4)

    def _load_data(self):
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("""
                SELECT td_sistol, td_diastol, nadi, alcohol_test,
                       kesimpulan_sistol, kesimpulan_diastol, kesimpulan_nadi, kesimpulan_alkohol,
                       keputusan, rekomendasi, keterangan,
                       avg_reaction, fatigue_level,
                       keluhan, diagnosis_awal, tindakan_medis, pemberian_obat
                FROM results WHERE id=?""", (self.db_id,))
            row = c.fetchone()
            conn.close()
            if row:
                if row[11] is not None:
                    self.var_rt.set(f"{row[11]:.1f} ms")
                if row[12]:
                    self.var_ft.set(str(row[12]))
                
                mapping = {
                    "td_sistol": 0, "td_diastol": 1, "nadi": 2, "alcohol_test": 3,
                    "kesimpulan_sistol": 4, "kesimpulan_diastol": 5, "kesimpulan_nadi": 6, 
                    "kesimpulan_alkohol": 7, "keputusan": 8, "rekomendasi": 9
                }
                for k, idx in mapping.items():
                    if k in self._vars and row[idx] is not None:
                        if k == "alcohol_test":
                            try: self._vars[k].set(f"{float(row[idx]):.2f}")
                            except: self._vars[k].set(str(row[idx]))
                        else:
                            self._vars[k].set(str(row[idx]))
                
                if row[10]:
                    self.txt_ket.insert("1.0", str(row[10]))
                    
                text_mapping = {}
                for k, idx in text_mapping.items():
                    if k in self._text_vars and row[idx] is not None:
                        self._text_vars[k].insert("1.0", str(row[idx]))
        except Exception as e:
            logging.error(f"Error loading medis data: {e}")

    def _auto_calc(self):
        """Auto-hitung kesimpulan dari nilai numerik yang diinput."""
        try:
            if "td_sistol" in self._vars:
                val = self._vars["td_sistol"].get().strip()
                if val:
                    try:
                        sis = float(val)
                        if   sis < 100: self._vars["kesimpulan_sistol"].set("Hipotensi")
                        elif sis < 140: self._vars["kesimpulan_sistol"].set("Normal")
                        elif sis < 160: self._vars["kesimpulan_sistol"].set("Hypertensi Ringan")
                        elif sis < 180: self._vars["kesimpulan_sistol"].set("Hypertensi Sedang")
                        else:           self._vars["kesimpulan_sistol"].set("Hypertensi Berat")
                    except:
                        pass
                else:
                    self._vars["kesimpulan_sistol"].set("")

            if "td_diastol" in self._vars:
                val = self._vars["td_diastol"].get().strip()
                if val:
                    try:
                        dia = float(val)
                        if   dia < 70:  self._vars["kesimpulan_diastol"].set("Hipotensi")
                        elif dia < 90:  self._vars["kesimpulan_diastol"].set("Normal")
                        elif dia < 100: self._vars["kesimpulan_diastol"].set("Hypertensi Ringan")
                        elif dia < 120: self._vars["kesimpulan_diastol"].set("Hypertensi Sedang")
                        else:           self._vars["kesimpulan_diastol"].set("Hypertensi Berat")
                    except:
                        pass
                else:
                    self._vars["kesimpulan_diastol"].set("")

            if "nadi" in self._vars:
                val = self._vars["nadi"].get().strip()
                if val:
                    try:
                        nad = float(val)
                        if   nad < 60:  self._vars["kesimpulan_nadi"].set("Bradikardia")
                        elif nad < 100: self._vars["kesimpulan_nadi"].set("Normal")
                        elif nad < 110: self._vars["kesimpulan_nadi"].set("Takikardia Ringan")
                        elif nad < 120: self._vars["kesimpulan_nadi"].set("Takikardia Sedang")
                        else:           self._vars["kesimpulan_nadi"].set("Takikardia Berat")
                    except:
                        pass
                else:
                    self._vars["kesimpulan_nadi"].set("")

            if "alcohol_test" in self._vars:
                val = self._vars["alcohol_test"].get().strip()
                if val:
                    try:
                        if float(val) < 0.08:
                            self._vars["kesimpulan_alkohol"].set("Lolos")
                        else:
                            self._vars["kesimpulan_alkohol"].set("Tidak Lolos")
                    except:
                        pass
                else:
                    self._vars["kesimpulan_alkohol"].set("")
        except Exception:
            pass

    def _save(self):
        try:
            if "alcohol_test" in self._vars:
                alc_val = self._vars["alcohol_test"].get().strip()
                if alc_val:
                    try: alc_val = f"{float(alc_val):.2f}"
                    except: pass
                    self._vars["alcohol_test"].set(alc_val)

            def get_var(k): return self._vars[k].get() if k in self._vars else None
            def get_txt(k): return self._text_vars[k].get("1.0", "end-1c").strip() if k in self._text_vars else None

            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("""
                    UPDATE results SET
                        td_sistol=COALESCE(?, td_sistol), 
                        td_diastol=COALESCE(?, td_diastol), 
                        nadi=COALESCE(?, nadi), 
                        alcohol_test=COALESCE(?, alcohol_test),
                        kesimpulan_sistol=COALESCE(?, kesimpulan_sistol), 
                        kesimpulan_diastol=COALESCE(?, kesimpulan_diastol), 
                        kesimpulan_nadi=COALESCE(?, kesimpulan_nadi), 
                        kesimpulan_alkohol=COALESCE(?, kesimpulan_alkohol),
                        keputusan=COALESCE(?, keputusan), 
                        rekomendasi=COALESCE(?, rekomendasi), 
                        keterangan=?,
                        keluhan=COALESCE(?, keluhan),
                        diagnosis_awal=COALESCE(?, diagnosis_awal),
                        tindakan_medis=COALESCE(?, tindakan_medis),
                        pemberian_obat=COALESCE(?, pemberian_obat)
                    WHERE id=?
                """, (
                    get_var("td_sistol") or None,
                    get_var("td_diastol") or None,
                    get_var("nadi") or None,
                    get_var("alcohol_test") or None,
                    get_var("kesimpulan_sistol"),
                    get_var("kesimpulan_diastol"),
                    get_var("kesimpulan_nadi"),
                    get_var("kesimpulan_alkohol"),
                    get_var("keputusan"),
                    get_var("rekomendasi"),
                    self.txt_ket.get("1.0", "end-1c").strip(),
                    get_txt("keluhan"),
                    get_txt("diagnosis_awal"),
                    get_txt("tindakan_medis"),
                    get_txt("pemberian_obat"),
                    self.db_id
                ))
                conn.commit()

            data_medis = {
                "td_sistol":           get_var("td_sistol") or "",
                "td_diastol":          get_var("td_diastol") or "",
                "nadi":                get_var("nadi") or "",
                "alcohol_test":        get_var("alcohol_test") or "",
                "kesimpulan_sistol":   get_var("kesimpulan_sistol") or "",
                "kesimpulan_diastol":  get_var("kesimpulan_diastol") or "",
                "kesimpulan_nadi":     get_var("kesimpulan_nadi") or "",
                "kesimpulan_alkohol":  get_var("kesimpulan_alkohol") or "",
                "keputusan":           get_var("keputusan") or "",
                "rekomendasi":         get_var("rekomendasi") or "",
                "keterangan":          self.txt_ket.get("1.0", "end-1c").strip(),
                "keluhan":             get_txt("keluhan") or "",
                "diagnosis_awal":      get_txt("diagnosis_awal") or "",
                "tindakan_medis":      get_txt("tindakan_medis") or "",
                "pemberian_obat":      get_txt("pemberian_obat") or ""
            }
            update_excel_medis(self.db_id, data_medis)
            logging.info(f"Medis saved for id={self.db_id}")
            self.on_save()
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

class TestMonitoringDialog(tk.Toplevel):
    def __init__(self, parent, on_close, on_open_medis=None):
        super().__init__(parent)
        self.title("Pengujian Fatigue Sedang Berlangsung")
        self.geometry("520x360")
        self.resizable(False, False)
        self.configure(bg=CLR["bg_panel"])
        self.on_close       = on_close
        self.on_open_medis  = on_open_medis
        self._result_shown  = False

        # ── Header ──────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=CLR["accent"], height=6)
        hdr.pack(fill="x")

        tk.Label(self, text="🧪 Pengujian FAST Aktif",
                 font=("Segoe UI", 13, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(16, 2))
        tk.Label(self, text="Mohon tekan tombol pada modul FAST sesuai instruksi.",
                 font=("Segoe UI", 9), fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack()

        # ── Status label ────────────────────────────────────────────
        self.lbl_status = tk.Label(
            self, text="⏸  Menunggu Tes…",
            font=("Segoe UI", 15, "bold"), fg=CLR["accent"], bg=CLR["bg_panel"],
            wraplength=470)
        self.lbl_status.pack(pady=(14, 4))

        # ── Dot legend ──────────────────────────────────────────────
        lf = tk.Frame(self, bg=CLR["bg_panel"])
        lf.pack()
        tk.Label(lf, text="⬤ Percobaan (T1-T2)",
                 font=("Segoe UI", 9), fg=CLR["warning"], bg=CLR["bg_panel"]).pack(side="left", padx=8)
        tk.Label(lf, text="⬤ Uji Asli (T3-T8)",
                 font=("Segoe UI", 9), fg=CLR["accent"], bg=CLR["bg_panel"]).pack(side="left", padx=8)

        # ── Dots ────────────────────────────────────────────────────
        self.dot_frame = tk.Frame(self, bg=CLR["bg_panel"])
        self.dot_frame.pack(pady=10)
        self.dots = []
        for i in range(8):
            lbl = tk.Label(self.dot_frame, text="○",
                           font=("Segoe UI", 22),
                           fg=CLR["separator"], bg=CLR["bg_panel"])
            lbl.pack(side="left", padx=4)
            self.dots.append(lbl)

        # ── Result panel (hidden awalnya) ────────────────────────────
        self.result_frame = tk.Frame(self, bg=CLR["bg_dark"],
                                     bd=0, relief="flat")
        # result_frame tidak di-pack dulu

        # ── Footer buttons ───────────────────────────────────────────
        self.btn_frame = tk.Frame(self, bg=CLR["bg_panel"])
        self.btn_frame.pack(pady=(8, 16))

        self.btn_medis = HoverButton(
            self.btn_frame, text="🩺 Input Medis",
            font=("Segoe UI", 10, "bold"),
            bg=CLR["accent2"], fg="white",
            relief="flat", cursor="hand2",
            padx=14, pady=6,
            state="disabled",
            command=self._go_medis)
        self.btn_medis.pack(side="left", padx=6)
        Tooltip(self.btn_medis, "Buka form input data medis setelah tes selesai")

        HoverButton(self.btn_frame, text="✕ Tutup",
                    font=("Segoe UI", 10), fg=CLR["text_muted"],
                    bg=CLR["bg_panel"], relief="flat", cursor="hand2",
                    padx=14, pady=6, command=self._on_closing).pack(side="left", padx=6)

        self._countdown_var = tk.StringVar(value="")
        self._countdown_lbl = tk.Label(self, textvariable=self._countdown_var,
                                       font=("Segoe UI", 9, "italic"),
                                       fg=CLR["text_muted"], bg=CLR["bg_panel"])
        self._countdown_lbl.pack()

        self.lift()
        self.focus_set()
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def set_status(self, text, color):
        self.lbl_status.configure(text=text, fg=color)

    def update_dot(self, idx):
        if 1 <= idx <= 8:
            self.dots[idx-1].configure(
                text="●",
                fg=CLR["warning"] if idx <= 2 else CLR["accent"])
            self.dot_frame.update_idletasks()

    def show_result(self, nama, avg_ms, fatigue_label):
        """Tampilkan panel hasil di dalam window ini setelah tes selesai."""
        if self._result_shown:
            return
        self._result_shown = True
        self.attributes("-topmost", False)
        self.geometry("520x480")
        self.title("Hasil Pengujian FAST")

        # Warna sesuai level fatigue
        color_map = {
            "Normal":         ("#1E8449", "#D5F5E3"),
            "Fatigue Ringan": ("#9A7D0A", "#FCF3CF"),
            "Fatigue Sedang": ("#A04000", "#FAE5D3"),
            "Fatigue Berat":  ("#922B21", "#FADBD8"),
        }
        fg_col, bg_col = color_map.get(fatigue_label, (CLR["text_main"], CLR["bg_dark"]))

        self.result_frame.configure(bg=bg_col)
        self.result_frame.pack(fill="x", padx=20, pady=(4, 10))

        tk.Label(self.result_frame,
                 text="✅  TES SELESAI",
                 font=("Segoe UI", 11, "bold"),
                 fg=fg_col, bg=bg_col).pack(pady=(10, 2))

        tk.Label(self.result_frame,
                 text=f"{nama}",
                 font=("Segoe UI", 10), fg=fg_col, bg=bg_col).pack()

        tk.Label(self.result_frame,
                 text=f"Rata-rata RT: {avg_ms:.1f} ms",
                 font=("Segoe UI", 14, "bold"), fg=fg_col, bg=bg_col).pack(pady=2)

        tk.Label(self.result_frame,
                 text=f"Tingkat Fatigue: {fatigue_label}",
                 font=("Segoe UI", 13, "bold"), fg=fg_col, bg=bg_col).pack(pady=(0, 10))

        # Aktifkan tombol Input Medis
        self.btn_medis.configure(state="normal")

        # Countdown 4 detik lalu auto-open medis
        self._start_countdown(4)

    def _start_countdown(self, seconds: int):
        """Countdown visual sebelum auto-open form medis."""
        if seconds <= 0:
            self._countdown_var.set("")
            self._go_medis()
            return
        self._countdown_var.set(f"Form medis otomatis terbuka dalam {seconds} detik…")
        self._countdown_job = self.after(1000, lambda: self._start_countdown(seconds - 1))

    def _go_medis(self):
        # Cancel countdown jika user klik manual
        if hasattr(self, "_countdown_job"):
            try:
                self.after_cancel(self._countdown_job)
            except Exception:
                pass
        if self.on_open_medis:
            self.on_open_medis()
        self._on_closing()

    def _on_closing(self):
        if hasattr(self, "_countdown_job"):
            try:
                self.after_cancel(self._countdown_job)
            except Exception:
                pass
        self.on_close()
        self.destroy()

# ==========================================
# MAIN APP
# ==========================================
class RatigApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{CONFIG.get('app_name', 'FAST')}")
        # Get screen width and height to calculate safe default window size
        try:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Capped at 1280 width and 700 height for standard desktop layouts
            default_width = min(1280, int(screen_width * 0.95))
            default_height = min(700, int(screen_height * 0.9))
            
            if default_width < 1024: default_width = 1024
            if default_height < 600: default_height = 600
            
            self.root.geometry(f"{default_width}x{default_height}")
        except Exception:
            self.root.geometry("1280x700")
            
        self.root.minsize(1024, 600)
        
        # Maximize window by default on startup (auto-fit screen)
        try:
            self.root.state('zoomed')
        except Exception:
            pass
        self.root.configure(bg=CLR["bg_dark"])

        # Window icon — load dari assets/ratig.ico jika ada
        _ico = os.path.join(ASSETS_DIR, "ratig.ico")
        if os.path.exists(_ico):
            try:
                self.root.iconbitmap(_ico)
            except Exception:
                pass

        self._apply_theme()

        # ── State vars ──────────────────────────────────────────────
        self.active_test_window = None
        self.current_user       = tk.StringVar(value="Anonim")
        self.current_nik        = tk.StringVar(value="")
        self.current_usia       = tk.StringVar(value="")
        self.current_jabatan    = tk.StringVar(value="Staff")
        # FIX: current_gender was referenced throughout but never created
        self.current_gender     = tk.StringVar(value="")
        self.current_divisi     = tk.StringVar(value="")
        self.current_contractor = tk.StringVar(value="")
        self.current_plan       = tk.StringVar(value="")

        self.selected_port  = tk.StringVar()
        self.conn_status    = tk.StringVar(value="Disconnected")
        self.status_bar_var = tk.StringVar(value="Siap.")
        self.search_var     = tk.StringVar()
        self.filter_from    = tk.StringVar(value="")
        self.filter_to      = tk.StringVar(value="")
        self.filter_jabatan = tk.StringVar(value="")
        self.filter_fatigue = tk.StringVar(value="Semua")
        # FIX: initialise filter_keputusan early so _get_filtered_df never raises AttributeError
        self.filter_keputusan = tk.StringVar(value="Semua")

        self._vcmd_num = (root.register(
            lambda v: v == "" or v.replace(".", "", 1).isdigit()), "%P")

        # progress dots state
        self._trial_vals  = [0.0] * 8
        self._trial_count = 0

        # layout
        self._build_layout()
        self._toggle_mode()
        self._update_nik_list()
        self.refresh_dashboard()

        # serial thread
        self.listener_running = True
        self.serial_conn      = None
        self._serial_queue    = queue.Queue()  # thread-safe queue untuk data dari serial
        threading.Thread(target=self._serial_loop,    daemon=True).start()
        threading.Thread(target=self._serial_processor, daemon=True).start()

        # auto backup hourly loop
        threading.Thread(target=auto_backup_loop, daemon=True).start()

        # auto refresh loop
        self._auto_refresh_loop()

        # global barcode / ID-card key scanner
        self._scan_buffer   = ""
        self._last_key_time = 0
        self.root.bind("<Key>", self._global_key_handler)

        # Konfirmasi sebelum tutup jika tes sedang berjalan
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

    # ─────────────────────────────────────────────────────────────────
    # THEME
    # ─────────────────────────────────────────────────────────────────
    def _apply_theme(self):
        s = ttk.Style()
        if "clam" in s.theme_names():
            s.theme_use("clam")
        s.configure("TNotebook", background=CLR["bg_dark"], borderwidth=0)
        s.configure("TNotebook.Tab",
                    background=CLR["bg_panel"],
                    foreground=CLR["text_muted"],
                    font=("Segoe UI", 10, "bold"),
                    padding=[14, 8], borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", CLR["bg_card"])],
              foreground=[("selected", CLR["accent"])])
        s.configure("Treeview",
                    background=CLR["bg_panel"],
                    foreground=CLR["text_main"],
                    fieldbackground=CLR["bg_panel"],
                    rowheight=26,
                    font=("Segoe UI", 9))
        s.configure("Treeview.Heading",
                    background=CLR["bg_card"],
                    foreground=CLR["accent"],
                    font=("Segoe UI", 9, "bold"),
                    relief="flat")
        s.map("Treeview",
              background=[("selected", CLR["accent"])],
              foreground=[("selected", "white")])
        s.configure("TScrollbar",
                    background=CLR["bg_card"],
                    troughcolor=CLR["bg_panel"],
                    borderwidth=0,
                    arrowcolor=CLR["text_muted"])
        s.configure("TCombobox",
                    fieldbackground=CLR["bg_card"],
                    background=CLR["bg_card"],
                    foreground=CLR["text_main"],
                    selectbackground=CLR["accent"],
                    font=("Segoe UI", 9),
                    postoffset=(0, 0, 200, 0))
        s.configure("Horizontal.TProgressbar",
                    troughcolor=CLR["bg_card"],
                    background=CLR["accent"],
                    borderwidth=0, thickness=10)

    # ─────────────────────────────────────────────────────────────────
    # LAYOUT
    # ─────────────────────────────────────────────────────────────────
    def _build_layout(self):
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()

    # ── SIDEBAR ──────────────────────────────────────────────────────
    def _build_sidebar(self):
        sb = tk.Frame(self.root, bg=CLR["bg_sidebar"], width=210)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.columnconfigure(0, weight=1)

        self.sidebar_app_name_label = tk.Label(sb, text=CONFIG.get("app_name", "FAST"),
                 font=("Segoe UI", 18, "bold"),
                 fg="white", bg=CLR["bg_sidebar"])
        self.sidebar_app_name_label.pack(pady=(10, 1))
        tk.Label(sb, text="Dashboard",
                 font=("Segoe UI", 8),
                 fg=CLR["sidebar_txt"],
                 bg=CLR["bg_sidebar"]).pack(pady=(0, 1))
        company = CONFIG.get("company_name", "")
        if company:
            tk.Label(sb, text=company[:28],
                     font=("Segoe UI", 7, "italic"),
                     fg=CLR["sidebar_sel"],
                     bg=CLR["bg_sidebar"],
                     wraplength=180).pack(pady=(0, 6))
        else:
            tk.Frame(sb, height=6, bg=CLR["bg_sidebar"]).pack()

        tk.Frame(sb, bg=CLR["sidebar_sel"], height=1).pack(fill="x", padx=16, pady=4)

        tk.Label(sb, text="FORM PENGUJIAN",
                 font=("Segoe UI", 8, "bold"),
                 fg=CLR["sidebar_txt"],
                 bg=CLR["bg_sidebar"]).pack(anchor="w", padx=16, pady=(4, 2))

        # Compact form frame with grid layout for side-by-side labels and entries
        form_frame = tk.Frame(sb, bg=CLR["bg_sidebar"])
        form_frame.pack(fill="x", padx=10, pady=2)
        form_frame.columnconfigure(1, weight=1)

        def add_form_row(row_idx, label_text, widget):
            lbl = tk.Label(form_frame, text=label_text, font=("Segoe UI", 8, "bold"),
                           fg=CLR["sidebar_txt"], bg=CLR["bg_sidebar"], anchor="w")
            lbl.grid(row=row_idx, column=0, padx=(6, 4), pady=2, sticky="w")
            widget.grid(row=row_idx, column=1, padx=(4, 6), pady=2, sticky="ew")

        # NIK Combobox (Menampilkan NIK - Nama)
        self.nik_cb = ttk.Combobox(form_frame, textvariable=self.current_nik,
                                   font=("Segoe UI", 9), width=15)
        self.nik_cb.bind("<<ComboboxSelected>>", self._autofill)
        self.nik_cb.bind("<KeyRelease>",         self._autofill)
        self.nik_cb.bind("<Return>",             self._autofill)
        self.nik_cb.bind("<FocusOut>",           self._autofill)
        add_form_row(0, "NIK", self.nik_cb)

        # Nama
        e_nama = tk.Entry(form_frame, textvariable=self.current_user, font=("Segoe UI", 9),
                          width=15, bg="#3D5166", fg="white",
                          insertbackground="white", relief="flat", bd=3)
        add_form_row(1, "Nama", e_nama)

        # Usia
        e_usia = tk.Entry(form_frame, textvariable=self.current_usia, font=("Segoe UI", 9),
                          width=15, bg="#3D5166", fg="white",
                          insertbackground="white", relief="flat", bd=3,
                          validate="key", validatecommand=self._vcmd_num)
        add_form_row(2, "Usia", e_usia)

        # Jenis Kelamin
        cb_jk = ttk.Combobox(form_frame, textvariable=self.current_gender, values=["Laki-laki", "Perempuan"],
                             font=("Segoe UI", 9), width=15, state="readonly")
        add_form_row(3, "Jenis Kel.", cb_jk)

        # Jabatan
        e_jabatan = tk.Entry(form_frame, textvariable=self.current_jabatan, font=("Segoe UI", 9),
                             width=15, bg="#3D5166", fg="white",
                             insertbackground="white", relief="flat", bd=3)
        add_form_row(4, "Jabatan", e_jabatan)

        tk.Frame(sb, bg=CLR["sidebar_sel"], height=1).pack(fill="x", padx=16, pady=4)

        HoverButton(sb, text="▶ Lakukan Test Fatigue",
                    command=self._open_test_dialog,
                    font=("Segoe UI", 11, "bold"),
                    bg=CLR["success"], fg="white",
                    relief="flat", cursor="hand2",
                    pady=6).pack(fill="x", padx=16, pady=(0, 4))

        tk.Frame(sb, bg=CLR["sidebar_sel"], height=1).pack(fill="x", padx=16, pady=4)

        tk.Label(sb, text="KONEKSI ESP32",
                 font=("Segoe UI", 8, "bold"),
                 fg=CLR["sidebar_txt"],
                 bg=CLR["bg_sidebar"]).pack(anchor="w", padx=16, pady=(2, 2))
        self.port_combo = ttk.Combobox(sb, textvariable=self.selected_port,
                                       font=("Segoe UI", 9), width=20)
        self.port_combo.pack(fill="x", padx=16, pady=(0, 2))
        self.port_combo.bind("<<ComboboxSelected>>", self._force_reconnect)
        self.refresh_ports()
        
        HoverButton(sb, text="↻ Refresh Port", command=self.refresh_ports,
                    font=("Segoe UI", 8), bg="#3D5166", fg="white",
                    relief="flat", cursor="hand2",
                    pady=2).pack(fill="x", padx=16, pady=(0, 2))
                    
        self._conn_label = tk.Label(sb, textvariable=self.conn_status,
                                    font=("Segoe UI", 8, "bold"),
                                    fg=CLR["danger"],
                                    bg=CLR["bg_sidebar"],
                                    wraplength=180)
        self._conn_label.pack(padx=16, pady=1)

        tk.Frame(sb, bg=CLR["sidebar_sel"], height=1).pack(fill="x", padx=16, pady=4)

        tk.Label(sb, text="MODE APLIKASI",
                 font=("Segoe UI", 8, "bold"),
                 fg=CLR["sidebar_txt"],
                 bg=CLR["bg_sidebar"]).pack(anchor="w", padx=16, pady=(2, 2))
        self.app_mode_var = tk.StringVar(value=LOCKED_MODE or "SHE Officer")
        mode_cb = ttk.Combobox(sb, textvariable=self.app_mode_var, values=["SHE Officer", "Paramedis"],
                                font=("Segoe UI", 9),
                                state="disabled" if LOCKED_MODE else "readonly")
        mode_cb.pack(fill="x", padx=16, pady=(0, 2))
        mode_cb.bind("<<ComboboxSelected>>", self._toggle_mode)

        tk.Frame(sb, bg=CLR["sidebar_sel"], height=1).pack(fill="x", padx=16, pady=4)

        HoverButton(sb, text="📝 Input Manual",
                    command=self._open_manual_input,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    pady=4).pack(fill="x", padx=16, pady=2)

        self._status_lbl = tk.Label(sb, textvariable=self.status_bar_var,
                                    font=("Segoe UI", 7, "italic"),
                                    fg=CLR["sidebar_txt"],
                                    bg=CLR["bg_sidebar"],
                                    wraplength=185, justify="left")
        self._status_lbl.pack(side="bottom", padx=10, pady=8, anchor="w")

    # ── MAIN CONTENT ─────────────────────────────────────────────────
    def _toggle_mode(self, event=None):
        if not hasattr(self, 'notebook') or not hasattr(self, 'tab_medis_frame'): return

        if (event is not None and LOCKED_MODE is None
                and self.app_mode_var.get() == "Paramedis"):
            pwd = simpledialog.askstring(
                "Akses Mode Paramedis",
                "Masukkan PIN untuk membuka Mode Paramedis:",
                show="*", parent=self.root)
            if not _verify_paramedis_pin(pwd):
                if pwd is not None:
                    messagebox.showerror("Ditolak", "PIN salah. Tetap di Mode SHE Officer.")
                self.app_mode_var.set("SHE Officer")
                self.notebook.tab(self.tab_medis_frame, state="hidden")
                return

        if self.app_mode_var.get() == "SHE Officer":
            self.notebook.tab(self.tab_medis_frame, state="hidden")
        else:
            self.notebook.tab(self.tab_medis_frame, state="normal")

        if hasattr(self, "_para_pin_widgets"):
            label, frame = self._para_pin_widgets
            if self.app_mode_var.get() == "Paramedis":
                label.grid()
                frame.grid()
            else:
                label.grid_remove()
                frame.grid_remove()

    def _build_main(self):
        right = tk.Frame(self.root, bg=CLR["bg_dark"])
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        self._build_notebook(right)

    def _open_test_dialog(self):
        nik = self.current_nik.get().strip().upper()
        if not nik:
            messagebox.showwarning("Ditolak",
                "NIK kosong!\nMasukkan NIK atau scan kartu terlebih dahulu.")
            return

        try:
            usia_val = int(self.current_usia.get())
            thr = CONFIG.get("age_warning_threshold", 60)
            if usia_val >= thr:
                if not messagebox.askyesno(
                    "Peringatan Usia",
                    f"Karyawan ini berusia {usia_val} tahun (Batas peringatan: {thr} tahun).\n\n"
                    "Mohon pastikan kondisi fisik karyawan memadai untuk dites.\n"
                    "Tetap lanjutkan pengujian?"):
                    return
        except Exception:
            pass

        # Validasi: NIK harus terdaftar di master karyawan atau kontraktor
        try:
            conn = sqlite3.connect(DB_PATH)
            in_kar = conn.execute(
                "SELECT 1 FROM master_karyawan WHERE nik=? LIMIT 1", (nik,)).fetchone()
            in_kon = conn.execute(
                "SELECT 1 FROM master_kontraktor WHERE nik=? LIMIT 1", (nik,)).fetchone()
            conn.close()
            if not in_kar and not in_kon:
                messagebox.showerror(
                    "NIK Tidak Terdaftar",
                    f"NIK '{nik}' tidak ditemukan di master Karyawan maupun Kontraktor.\n\n"
                    "Daftarkan karyawan terlebih dahulu di tab Master Karyawan / Master Kontraktor "
                    "sebelum melakukan pengujian.")
                return
        except Exception as e:
            logging.error(f"Master NIK check: {e}")

        today = datetime.datetime.now().strftime("%Y-%m-%d")
        limit = CONFIG.get("daily_test_limit", 3)
        try:
            conn = sqlite3.connect(DB_PATH)
            tests_today = conn.execute(
                "SELECT COUNT(*) FROM results WHERE nik=? AND tanggal LIKE ?",
                (nik, today + "%")).fetchone()[0]
            conn.close()
            if tests_today >= limit:
                messagebox.showwarning(
                    "Tes Ditolak",
                    f"NIK {nik} sudah dites {tests_today}x hari ini!\n"
                    f"Berdasarkan kebijakan, satu orang hanya boleh tes maksimal {limit}x per hari.")
                return
        except Exception as e:
            logging.error(f"Duplicate check error: {e}")

        if hasattr(self, 'active_test_window') and self.active_test_window is not None:
            return

        def _open_medis_callback():
            """Dipanggil dari tombol 'Input Medis' dalam dialog hasil."""
            if hasattr(self, '_pending_medis_id') and self._pending_medis_id:
                self._open_medis_dialog_by_id(self._pending_medis_id)

        self._pending_medis_id = None
        self.active_test_window = TestMonitoringDialog(
            self.root,
            on_close=self._on_test_dialog_close,
            on_open_medis=_open_medis_callback)

    def _on_test_dialog_close(self):
        self.active_test_window = None

    def _build_notebook(self, parent):
        nb_frame = tk.Frame(parent, bg=CLR["bg_dark"])
        nb_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(6, 10))
        nb_frame.columnconfigure(0, weight=1)
        nb_frame.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(nb_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self._tab_data()
        self._tab_medis()
        self._tab_stats()
        self._tab_filter()
        self._tab_master()
        self._tab_master_kontraktor()
        self._tab_settings()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._bind_context_menus()

    # ─────────────────────────────────────────────────────────────────
    # TAB 1 — DATA & PENGUJIAN
    # ─────────────────────────────────────────────────────────────────
    def _tab_data(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  📋 Data & Pengujian  ")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ctrl = tk.Frame(tab, bg=CLR["bg_panel"], pady=8, padx=10)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(6, 0))

        tk.Label(ctrl, text="Cari Nama / NIK:", font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left")
        e_search = tk.Entry(ctrl, textvariable=self.search_var, font=("Segoe UI", 9),
                 width=22, bg=CLR["bg_card"], fg=CLR["text_main"],
                 insertbackground=CLR["accent"],
                 relief="flat", bd=4)
        e_search.pack(side="left", padx=6)
        e_search.bind("<Return>", lambda e: self._search())
        HoverButton(ctrl, text="Cari", command=self._search,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left")
        HoverButton(ctrl, text="Tampilkan Semua",
                    command=lambda: self.update_table(reset=True),
                    font=("Segoe UI", 9),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=6)
        HoverButton(ctrl, text="🩺 Input Medis",
                    command=self._open_medis_dialog,
                    font=("Segoe UI", 9, "bold"),
                    bg="#8E44AD", fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=6)
        HoverButton(ctrl, text="📥 Import Hasil Tes",
                    command=self.import_data,
                    font=("Segoe UI", 9, "bold"),
                    bg="#F39C12", fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=6)
        HoverButton(ctrl, text="⬇ Ekspor Excel",
                    command=self.export_excel,
                    font=("Segoe UI", 9, "bold"),
                    bg="#27AE60", fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=6)
        HoverButton(ctrl, text="🗑 Hapus Semua Data",
                    command=self._reset_data,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["danger"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="right", padx=6)

        tf = self._make_tree_frame(tab, row=1)
        # Columns: ID(0), Tanggal(1), Nama(2), NIK(3), Usia(4),
        #          Bagian(5), Jabatan(6), T1..T6(7-12), Status(13), Fatigue(14)
        cols   = ("ID", "Tanggal", "Nama", "NIK", "Usia", "Jenis Kelamin",
                  "Jabatan / Bagian", "Mode", "T1", "T2", "T3", "T4", "T5", "T6",
                  "Status", "Fatigue")
        widths = [40, 130, 110, 90, 40, 90, 120, 70, 45, 45, 45, 45, 45, 45, 70, 100]
        self.tree = self._make_tree(tf, cols, widths, sortable=True)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

    # ─────────────────────────────────────────────────────────────────
    # TAB 2 — DATA MEDIS
    # ─────────────────────────────────────────────────────────────────
    def _tab_medis(self):
        self.tab_medis_frame = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(self.tab_medis_frame, text="  🏥 Data Medis  ")
        tab = self.tab_medis_frame
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ctrl = tk.Frame(tab, bg=CLR["bg_panel"], pady=8, padx=10)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(6, 0))

        tk.Label(ctrl, text="Pencarian Medis:", font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left")
        self.search_medis_var = tk.StringVar()
        e_medis = tk.Entry(ctrl, textvariable=self.search_medis_var, font=("Segoe UI", 9),
                 width=22, bg=CLR["bg_card"], fg=CLR["text_main"],
                 insertbackground=CLR["accent"],
                 relief="flat", bd=4)
        e_medis.pack(side="left", padx=6)
        e_medis.bind("<Return>", lambda e: self.update_table_medis())
        HoverButton(ctrl, text="Cari",
                    command=lambda: self.update_table_medis(),
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left")
        HoverButton(ctrl, text="Tampilkan Semua",
                    command=lambda: self.update_table_medis(reset=True),
                    font=("Segoe UI", 9),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=6)
        HoverButton(ctrl, text="📥 Import Hasil Tes",
                    command=self.import_data,
                    font=("Segoe UI", 9, "bold"),
                    bg="#F39C12", fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="right", padx=6)
        HoverButton(ctrl, text="⬇ Ekspor Rekap Medis",
                    command=self.export_excel,
                    font=("Segoe UI", 9, "bold"),
                    bg="#2980B9", fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="right", padx=6)

        tf = self._make_tree_frame(tab, row=1)
        cols   = ("ID", "Tanggal", "Nama", "NIK", "Jenis Kelamin", "Jabatan / Bagian",
                  "RT (ms)", "Fatigue",
                  "Sistol", "Diastol", "Nadi", "Alkohol",
                  "Kesp. Sistol", "Kesp. Diastol", "Kesp. Nadi",
                  "Keputusan", "Rekomendasi",
                  "Keterangan")
        widths = [40, 130, 110, 90, 90, 120, 60, 100, 50, 50, 50, 60, 110, 110, 110, 140, 160, 200]
        self.tree_medis = self._make_tree(tf, cols, widths, sortable=True)
        self.tree_medis.bind("<Double-1>",
                             lambda e: self._open_medis_dialog(from_medis_tab=True))
        self.tree_medis.bind("<ButtonRelease-1>", self._on_medis_click)

        # Quick keterangan bar
        tab.rowconfigure(2, weight=0)
        ket_frame = tk.Frame(tab, bg=CLR["bg_panel"], pady=6, padx=10)
        ket_frame.grid(row=2, column=0, sticky="ew")
        tk.Label(ket_frame,
                 text="Keterangan Medis (Pilih baris lalu Edit & Enter):",
                 font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left")
        self.ket_medis_var = tk.StringVar()
        self.ket_medis_id  = None
        self.ket_entry = tk.Entry(ket_frame, textvariable=self.ket_medis_var,
                                  font=("Segoe UI", 9),
                                  bg=CLR["bg_card"], fg=CLR["text_main"],
                                  insertbackground=CLR["accent"],
                                  relief="flat", bd=4)
        self.ket_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.ket_entry.bind("<Return>", self._save_keterangan_cepat)

    # ─────────────────────────────────────────────────────────────────
    # TAB 3 — STATISTIK
    # ─────────────────────────────────────────────────────────────────
    def _tab_stats(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  📊 Statistik  ")
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)
        tab.rowconfigure(1, weight=1)
        tab.rowconfigure(2, weight=1)

        kpi = tk.Frame(tab, bg=CLR["bg_dark"])
        kpi.grid(row=0, column=0, columnspan=2,
                 sticky="ew", padx=10, pady=10)
        for i in range(5):
            kpi.columnconfigure(i, weight=1)
        self._kpi_total  = self._kpi(kpi, "Total Pengujian",  "0", 0)
        self._kpi_normal = self._kpi(kpi, "Normal",           "0", 1)
        self._kpi_ringan = self._kpi(kpi, "Fatigue Ringan",   "0", 2)
        self._kpi_sedang = self._kpi(kpi, "Fatigue Sedang",   "0", 3)
        self._kpi_berat  = self._kpi(kpi, "Fatigue Berat",    "0", 4)

        cL = tk.Frame(tab, bg=CLR["bg_panel"], padx=6, pady=6)
        cL.grid(row=1, column=0, sticky="nsew", padx=(10, 4), pady=(0, 4))
        cR = tk.Frame(tab, bg=CLR["bg_panel"], padx=6, pady=6)
        cR.grid(row=1, column=1, sticky="nsew", padx=(4, 10), pady=(0, 4))

        self.fig_daily = Figure(figsize=(5, 3.2), dpi=96, facecolor=CLR["bg_panel"])
        self.ax_daily  = self.fig_daily.add_subplot(111)
        self._sax(self.ax_daily)
        FigureCanvasTkAgg(self.fig_daily, master=cL).get_tk_widget().pack(
            fill="both", expand=True)

        self.fig_pie = Figure(figsize=(5, 3.2), dpi=96, facecolor=CLR["bg_panel"])
        self.ax_pie  = self.fig_pie.add_subplot(111)
        self._sax(self.ax_pie)
        FigureCanvasTkAgg(self.fig_pie, master=cR).get_tk_widget().pack(
            fill="both", expand=True)

        cD = tk.Frame(tab, bg=CLR["bg_panel"], padx=6, pady=6)
        cD.grid(row=2, column=0, columnspan=2,
                sticky="nsew", padx=10, pady=(0, 10))
        self.fig_divisi = Figure(figsize=(10, 2.8), dpi=96, facecolor=CLR["bg_panel"])
        self.ax_divisi  = self.fig_divisi.add_subplot(111)
        self._sax(self.ax_divisi)
        FigureCanvasTkAgg(self.fig_divisi, master=cD).get_tk_widget().pack(
            fill="both", expand=True)

    def _kpi(self, parent, label, val, col):
        card = tk.Frame(parent, bg=CLR["bg_card"], padx=14, pady=10)
        card.grid(row=0, column=col, sticky="ew", padx=4)
        tk.Label(card, text=label, font=("Segoe UI", 8),
                 fg=CLR["text_muted"], bg=CLR["bg_card"]).pack(anchor="w")
        lbl = tk.Label(card, text=val,
                       font=("Segoe UI", 17, "bold"),
                       fg=CLR["accent"], bg=CLR["bg_card"])
        lbl.pack(anchor="w")
        return lbl

    @staticmethod
    def _sax(ax):
        ax.set_facecolor(CLR["bg_panel"])
        for sp in ax.spines.values():
            sp.set_color(CLR["separator"])
        ax.tick_params(colors=CLR["text_muted"], labelsize=7)
        ax.title.set_color(CLR["text_main"])

    # ─────────────────────────────────────────────────────────────────
    # TAB 4 — FILTER & RIWAYAT
    # ─────────────────────────────────────────────────────────────────
    def _tab_filter(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  🔍 Filter & Riwayat  ")
        tab.columnconfigure(0, weight=1)
        # FIX: correct row weights — row 0 = filter bar1, 1 = filter bar2,
        #      2 = tree (expand), 3 = action buttons
        tab.rowconfigure(2, weight=1)

        # --- Filter bar 1 ---
        fctl = tk.Frame(tab, bg=CLR["bg_panel"], pady=10, padx=14)
        fctl.grid(row=0, column=0, sticky="ew", pady=(6, 0))

        def fl(t):
            return tk.Label(fctl, text=t, font=("Segoe UI", 9),
                            fg=CLR["text_muted"], bg=CLR["bg_panel"])

        def fe(v, w=13):
            ent = tk.Entry(fctl, textvariable=v, font=("Segoe UI", 9),
                            width=w, bg=CLR["bg_card"], fg=CLR["text_main"],
                            insertbackground=CLR["accent"],
                            relief="flat", bd=4)
            ent.bind("<Return>", lambda e: self._apply_filter())
            return ent

        fl("Dari (YYYY-MM-DD):").pack(side="left", padx=(0, 3))
        fe(self.filter_from).pack(side="left", padx=3)
        fl("Sampai:").pack(side="left", padx=(6, 3))
        fe(self.filter_to).pack(side="left", padx=3)
        fl("Jabatan:").pack(side="left", padx=(8, 3))
        fe(self.filter_jabatan, w=15).pack(side="left", padx=3)
        fl("Fatigue:").pack(side="left", padx=(8, 3))
        cb_fatigue = ttk.Combobox(fctl, textvariable=self.filter_fatigue,
                     values=["Semua", "Normal", "Fatigue Ringan",
                             "Fatigue Sedang", "Fatigue Berat"],
                     font=("Segoe UI", 9), width=13,
                     state="readonly")
        cb_fatigue.pack(side="left", padx=3)
        cb_fatigue.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())
        HoverButton(fctl, text="Terapkan",
                    command=self._apply_filter,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left", padx=7)
        HoverButton(fctl, text="Reset",
                    command=self._reset_filter,
                    font=("Segoe UI", 9),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", cursor="hand2",
                    padx=8, pady=3).pack(side="left")

        # --- Filter bar 2 (keputusan) ---
        fctl2 = tk.Frame(tab, bg=CLR["bg_panel"], pady=6, padx=14)
        fctl2.grid(row=1, column=0, sticky="ew")
        tk.Label(fctl2, text="Keputusan:", font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left", padx=(0, 3))
        cb_keputusan = ttk.Combobox(
            fctl2, textvariable=self.filter_keputusan,
            values=["Semua", "Diizinkan bekerja",
                    "Diizinkan bekerja, dengan catatan",
                    "Tidak diizinkan bekerja",
                    "Rujuk ke Faskes"],
            font=("Segoe UI", 9), width=35,
            state="readonly"
        )
        cb_keputusan.pack(side="left", padx=3)
        cb_keputusan.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())

        # --- Tree --- (row=2, expands)
        tf = self._make_tree_frame(tab, row=2)
        cols   = ("ID", "Tanggal", "Nama", "NIK", "Jenis Kelamin", "Jabatan",
                  "T1", "T2", "T3", "T4", "T5", "T6", "Avg",
                  "Total", "Sistol", "Diastol", "Nadi", "Fatigue", "Keputusan")
        widths = [40, 130, 110, 90, 90, 90, 40, 40, 40, 40, 40, 40, 45, 55, 55, 55, 45, 100, 140]
        self.tree_filter = self._make_tree(tf, cols, widths)
        self.tree_filter.bind("<Double-1>", self._on_filter_tree_double_click)

        # --- Action buttons --- (row=3)
        act = tk.Frame(tab, bg=CLR["bg_dark"])
        act.grid(row=3, column=0, sticky="ew", padx=6, pady=4)
        bk = dict(font=("Segoe UI", 9, "bold"), fg="white",
                  relief="flat", cursor="hand2", padx=12, pady=4)
        HoverButton(act, text="⬇ Excel (Filter)", command=self.export_excel,
                    bg="#27AE60", **bk).pack(side="left", padx=3)
        HoverButton(act, text="⬇ CSV (Filter)",   command=self.export_csv,
                    bg="#2980B9", **bk).pack(side="left", padx=3)
        HoverButton(act, text="⬇ PDF (Filter)",   command=self.export_pdf,
                    bg="#C0392B", **bk).pack(side="left", padx=3)
        HoverButton(act, text="⬆ Import Hasil Tes", command=self.import_data,
                    bg="#F39C12", **bk).pack(side="left", padx=3)
        HoverButton(act, text="🗑 Hapus Terpilih",
                    command=self._delete_selected,
                    bg=CLR["danger"], **bk).pack(side="right", padx=3)

    # ─────────────────────────────────────────────────────────────────
    # TAB 5 — MASTER KARYAWAN
    # ─────────────────────────────────────────────────────────────────
    def _tab_master(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  👥 Master Karyawan  ")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ctrl = tk.Frame(tab, bg=CLR["bg_panel"], pady=8, padx=10)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(6, 0))

        self._mk_search = tk.StringVar()
        tk.Label(ctrl, text="Cari:", font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left")
        e_mk = tk.Entry(ctrl, textvariable=self._mk_search, font=("Segoe UI", 9),
                 width=20, bg=CLR["bg_card"], fg=CLR["text_main"],
                 insertbackground=CLR["accent"],
                 relief="flat", bd=4)
        e_mk.pack(side="left", padx=6)
        e_mk.bind("<Return>", lambda e: self._mk_refresh())
        HoverButton(ctrl, text="Cari", command=self._mk_refresh,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left")

        bk = dict(font=("Segoe UI", 9, "bold"), fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=3)
        HoverButton(ctrl, text="➕ Tambah",     command=self._mk_add,
                    bg=CLR["success"], **bk).pack(side="left", padx=(10, 3))
        HoverButton(ctrl, text="✏️ Edit",       command=self._mk_edit,
                    bg=CLR["warning"], **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="🗑 Hapus",      command=self._mk_delete,
                    bg=CLR["danger"],  **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="🗑 Hapus Semua", command=self._mk_delete_all,
                    bg=CLR["danger"],  **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="📥 Import Excel",
                    command=self._import_master,
                    bg="#8E44AD", **bk).pack(side="right", padx=6)
        HoverButton(ctrl, text="⬇ Template Import",
                    command=self._download_template_master_karyawan,
                    bg="#2980B9", **bk).pack(side="right", padx=3)
        HoverButton(ctrl, text="⬆ Export Excel",
                    command=self._export_master,
                    bg="#27AE60", **bk).pack(side="right", padx=3)

        tf = self._make_tree_frame(tab, row=1)
        cols   = ("No", "Nama", "NIK", "Usia", "Jenis Kelamin", "Jabatan", "Plan/Devisi/Perusahaan")
        widths = [40, 140, 110, 50, 90, 100, 250]
        self.tree_master = self._make_tree(tf, cols, widths)
        self.tree_master.bind("<Double-1>", lambda e: self._select_master_for_test(self.tree_master))
        self._mk_refresh()

    # ─────────────────────────────────────────────────────────────────
    # TAB 6 — MASTER KONTRAKTOR
    # ─────────────────────────────────────────────────────────────────
    def _tab_master_kontraktor(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  👷 Master Kontraktor  ")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ctrl = tk.Frame(tab, bg=CLR["bg_panel"], pady=8, padx=10)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(6, 0))

        self._mkon_search = tk.StringVar()
        tk.Label(ctrl, text="Cari:", font=("Segoe UI", 9),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"]).pack(side="left")
        e_mkon = tk.Entry(ctrl, textvariable=self._mkon_search, font=("Segoe UI", 9),
                 width=20, bg=CLR["bg_card"], fg=CLR["text_main"],
                 insertbackground=CLR["accent"],
                 relief="flat", bd=4)
        e_mkon.pack(side="left", padx=6)
        e_mkon.bind("<Return>", lambda e: self._mkon_refresh())
        HoverButton(ctrl, text="Cari", command=self._mkon_refresh,
                    font=("Segoe UI", 9, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=10, pady=3).pack(side="left")

        bk = dict(font=("Segoe UI", 9, "bold"), fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=3)
        HoverButton(ctrl, text="➕ Tambah",     command=self._mkon_add,
                    bg=CLR["success"], **bk).pack(side="left", padx=(10, 3))
        HoverButton(ctrl, text="✏️ Edit",       command=self._mkon_edit,
                    bg=CLR["warning"], **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="🗑 Hapus",      command=self._mkon_delete,
                    bg=CLR["danger"],  **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="🗑 Hapus Semua", command=self._mkon_delete_all,
                    bg=CLR["danger"],  **bk).pack(side="left", padx=3)
        HoverButton(ctrl, text="📥 Import Excel",
                    command=self._import_master_kontraktor,
                    bg="#8E44AD", **bk).pack(side="right", padx=6)
        HoverButton(ctrl, text="⬇ Template Import",
                    command=self._download_template_master_kontraktor,
                    bg="#2980B9", **bk).pack(side="right", padx=3)
        HoverButton(ctrl, text="⬆ Export Excel",
                    command=self._export_master_kontraktor,
                    bg="#27AE60", **bk).pack(side="right", padx=3)

        tf = self._make_tree_frame(tab, row=1)
        cols   = ("No", "Nama", "NIK", "Usia", "Jenis Kelamin", "Jabatan", "Nama PT / Perusahaan")
        widths = [40, 140, 110, 50, 90, 100, 250]
        self.tree_master_kontraktor = self._make_tree(tf, cols, widths)
        self.tree_master_kontraktor.bind("<Double-1>", lambda e: self._select_master_for_test(self.tree_master_kontraktor))
        self._mkon_refresh()

    # ─────────────────────────────────────────────────────────────────
    # TAB 7 — PENGATURAN
    # ─────────────────────────────────────────────────────────────────
    def _tab_settings(self):
        tab = tk.Frame(self.notebook, bg=CLR["bg_dark"])
        self.notebook.add(tab, text="  ⚙ Pengaturan  ")

        inner = tk.Frame(tab, bg=CLR["bg_panel"], padx=24, pady=20)
        inner.pack(padx=20, pady=20, anchor="nw")

        self._cfg_baud    = tk.StringVar(value=str(CONFIG["baudrate"]))
        self._cfg_company = tk.StringVar(value=CONFIG.get("company_name", ""))
        self._cfg_created_by = tk.StringVar(value=CONFIG.get("created_by", "Marsellinus, Khilqi, Idris"))
        self._cfg_excel   = tk.StringVar(value=EXCEL_PATH)
        self._cfg_db      = tk.StringVar(value=DB_PATH)
        self._cfg_auto    = tk.BooleanVar(value=CONFIG["auto_connect"])
        self._cfg_backup  = tk.BooleanVar(value=CONFIG.get("auto_backup", True))
        self._cfg_app_name = tk.StringVar(value=CONFIG.get("app_name", "FAST"))
        self._cfg_age_thr = tk.StringVar(
            value=str(CONFIG.get("age_warning_threshold", 60)))
        self._cfg_limit_test = tk.StringVar(
            value=str(CONFIG.get("daily_test_limit", 3)))

        def row(r, lbl, var, ro=False):
            tk.Label(inner, text=lbl, font=("Segoe UI", 10),
                     fg=CLR["text_muted"], bg=CLR["bg_panel"],
                     width=26, anchor="w").grid(
                row=r, column=0, padx=(0, 10), pady=5, sticky="w")
            tk.Entry(inner, textvariable=var, font=("Segoe UI", 10),
                     width=40, bg="#F1F5F9", fg=CLR["text_main"],
                     insertbackground=CLR["accent"],
                     relief="flat", bd=0,
                     highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                     highlightthickness=1,
                     state="readonly" if ro else "normal").grid(
                row=r, column=1, pady=5)

        def chk(r, lbl, var):
            tk.Label(inner, text=lbl, font=("Segoe UI", 10),
                     fg=CLR["text_muted"], bg=CLR["bg_panel"],
                     width=26, anchor="w").grid(row=r, column=0, pady=5, sticky="w")
            tk.Checkbutton(inner, variable=var,
                           bg=CLR["bg_panel"],
                           activebackground=CLR["bg_panel"],
                           selectcolor=CLR["bg_card"],
                           fg=CLR["text_main"]).grid(row=r, column=1, sticky="w")

        row(0, "Nama Perusahaan",          self._cfg_company)
        row(1, "Dibuat Oleh (Developer)",   self._cfg_created_by)
        row(2, "Baud Rate",                self._cfg_baud)
        row(3, "Batas Usia Peringatan",    self._cfg_age_thr)
        row(4, "Batas Tes Karyawan/Hari",   self._cfg_limit_test)

        tk.Label(inner, text="PIN Hapus Data", font=("Segoe UI", 10),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"],
                 width=26, anchor="w").grid(
            row=4, column=0, padx=(0, 10), pady=5, sticky="w")

        pin_frame = tk.Frame(inner, bg=CLR["bg_panel"])
        pin_frame.grid(row=4, column=1, pady=5, sticky="w")

        tk.Entry(pin_frame, textvariable=tk.StringVar(value="********"), font=("Segoe UI", 10),
                 width=18, bg="#F1F5F9", fg=CLR["text_main"],
                 relief="flat", bd=0, state="readonly", show="*",
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                 highlightthickness=1).pack(side="left")

        HoverButton(pin_frame, text="Ubah PIN", command=self._change_pin_dialog,
                    font=("Segoe UI", 9, "bold"), bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2", padx=10, pady=2).pack(side="left", padx=(10, 0))

        HoverButton(pin_frame, text="Lupa PIN?", command=self._reset_pin_dialog,
                    font=("Segoe UI", 8, "underline"), bg=CLR["bg_panel"], fg=CLR["accent"],
                    relief="flat", cursor="hand2").pack(side="left", padx=(10, 0))

        para_pin_label = tk.Label(inner, text="PIN Mode Paramedis", font=("Segoe UI", 10),
                 fg=CLR["text_muted"], bg=CLR["bg_panel"],
                 width=26, anchor="w")
        para_pin_label.grid(
            row=5, column=0, padx=(0, 10), pady=5, sticky="w")

        para_pin_frame = tk.Frame(inner, bg=CLR["bg_panel"])
        para_pin_frame.grid(row=5, column=1, pady=5, sticky="w")
        self._para_pin_widgets = (para_pin_label, para_pin_frame)

        tk.Entry(para_pin_frame, textvariable=tk.StringVar(value="********"), font=("Segoe UI", 10),
                 width=18, bg="#F1F5F9", fg=CLR["text_main"],
                 relief="flat", bd=0, state="readonly", show="*",
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"],
                 highlightthickness=1).pack(side="left")

        HoverButton(para_pin_frame, text="Ubah PIN", command=self._change_paramedis_pin_dialog,
                    font=("Segoe UI", 9, "bold"), bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2", padx=10, pady=2).pack(side="left", padx=(10, 0))

        HoverButton(para_pin_frame, text="Lupa PIN?", command=self._reset_paramedis_pin_dialog,
                    font=("Segoe UI", 8, "underline"), bg=CLR["bg_panel"], fg=CLR["accent"],
                    relief="flat", cursor="hand2").pack(side="left", padx=(10, 0))

        row(6, "Path Excel",               self._cfg_excel, ro=True)
        row(7, "Path Database",            self._cfg_db,    ro=True)
        chk(8, "Auto-connect Serial",      self._cfg_auto)
        chk(9, "Auto-backup Harian",       self._cfg_backup)

        row(10, "Nama Aplikasi",           self._cfg_app_name)

        HoverButton(inner, text="Simpan Pengaturan",
                    command=self._save_settings,
                    font=("Segoe UI", 10, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2",
                    padx=16, pady=6
                    ).grid(row=10, column=0, columnspan=2, pady=14, sticky="w")

        tk.Label(inner,
                 text=f"Log: {LOG_FILE}  |  Backup folder: {BACKUP_DIR}",
                 font=("Segoe UI", 8), fg=CLR["text_muted"],
                 bg=CLR["bg_panel"]
                 ).grid(row=11, column=0, columnspan=2, sticky="w")

        tk.Frame(inner, bg=CLR["separator"], height=1).grid(
            row=12, column=0, columnspan=2, sticky="ew", pady=(16, 8))
        tk.Label(inner, text="Impor Data",
                 font=("Segoe UI", 11, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]
                 ).grid(row=13, column=0, columnspan=2, sticky="w", pady=(0, 8))
        HoverButton(inner, text="📥 Import Riwayat Tes (Excel/CSV)",
                    command=self.import_data,
                    font=("Segoe UI", 10, "bold"),
                    bg="#8E44AD", fg="white",
                    relief="flat", cursor="hand2",
                    padx=16, pady=6
                    ).grid(row=14, column=0, columnspan=2, sticky="w")

        tk.Frame(inner, bg=CLR["separator"], height=1).grid(
            row=15, column=0, columnspan=2, sticky="ew", pady=(16, 8))

        about_f = tk.Frame(inner, bg=CLR["bg_panel"])
        about_f.grid(row=16, column=0, columnspan=2, sticky="w")
        
        self.about_app_name_label = tk.Label(about_f, text=CONFIG.get("app_name", "FAST"), font=("Segoe UI", 16, "bold"),
                 fg=CLR["accent"], bg=CLR["bg_panel"])
        self.about_app_name_label.pack(anchor="w")
        self.about_app_title_label = tk.Label(about_f, text="",
                 font=("Segoe UI", 10), fg=CLR["text_muted"],
                 bg=CLR["bg_panel"])
        self.about_app_title_label.pack(anchor="w")
        self.about_created_by_label = tk.Label(about_f, text=f"Dibuat Oleh: {CONFIG.get('created_by', 'Marsellinus, Khilqi, Idris')}",
                 font=("Segoe UI", 9, "bold"), fg=CLR["accent2"],
                 bg=CLR["bg_panel"])
        self.about_created_by_label.pack(anchor="w", pady=(4, 8))

        btn_f = tk.Frame(about_f, bg=CLR["bg_panel"])
        btn_f.pack(anchor="w")
        HoverButton(btn_f, text="📂 Buka Folder Data", command=lambda: _open_file(DATA_DIR),
                    font=("Segoe UI", 8, "bold"), bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2", padx=8, pady=3).pack(side="left", padx=(0, 6))
        HoverButton(btn_f, text="📂 Buka Folder Backup", command=lambda: _open_file(BACKUP_DIR),
                    font=("Segoe UI", 8, "bold"), bg=CLR["accent2"], fg="white",
                    relief="flat", cursor="hand2", padx=8, pady=3).pack(side="left", padx=(0, 6))
        HoverButton(btn_f, text="📄 Buka Log", command=lambda: _open_file(LOG_FILE),
                    font=("Segoe UI", 8, "bold"), bg=CLR["warning"], fg="white",
                    relief="flat", cursor="hand2", padx=8, pady=3).pack(side="left")

    # ─────────────────────────────────────────────────────────────────
    # TREE HELPERS
    # ─────────────────────────────────────────────────────────────────
    def _make_tree_frame(self, parent, row):
        f = tk.Frame(parent, bg=CLR["bg_dark"])
        f.grid(row=row, column=0, sticky="nsew", pady=6, padx=6)
        f.columnconfigure(0, weight=1)
        f.rowconfigure(0, weight=1)
        return f

    def _make_tree(self, parent, cols, widths, sortable=False):
        tv = ttk.Treeview(parent, columns=cols, show="headings", height=14)
        tv.tag_configure("even",      background=CLR["row_even"],
                         foreground=CLR["text_main"])
        tv.tag_configure("odd",       background=CLR["row_odd"],
                         foreground=CLR["text_main"])
        tv.tag_configure("last_tested", background="#EBF5FB",
                         foreground=CLR["text_main"])
        for fl, bg in CLR["fatigue_colors"].items():
            safe_tag = "fl_" + fl.replace(" ", "_")
            tv.tag_configure(safe_tag,
                             background=bg if bg else CLR["row_even"],
                             foreground=CLR["fatigue_fg"].get(fl, CLR["text_main"]))
        for col, w in zip(cols, widths):
            if sortable:
                cmd = lambda c=col: self._sort_tree(tv, c, False)
                tv.heading(col, text=col, command=cmd)
            else:
                tv.heading(col, text=col)
            tv.column(col, width=w, anchor="center", minwidth=30)
        vsb = ttk.Scrollbar(parent, orient="vertical",   command=tv.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        return tv

    def _sort_tree(self, tv, col, rev):
        rows = [(tv.set(k, col), k) for k in tv.get_children("")]
        try:
            rows.sort(key=lambda t: float(t[0]), reverse=rev)
        except Exception:
            rows.sort(reverse=rev)
        for i, (_, k) in enumerate(rows):
            tv.move(k, "", i)
        tv.heading(col, command=lambda: self._sort_tree(tv, col, not rev))

    def _populate_tree(self, tree, df, highlight_first=False, preserve_selection=False):
        sel_id = None
        if preserve_selection:
            sel = tree.selection()
            if sel:
                try: sel_id = str(tree.item(sel[0])["values"][0])
                except: pass

        for item in tree.get_children():
            tree.delete(item)
            
        self._last_iid = None
        restore_iid = None
        
        for count, row in enumerate(df.itertuples(index=False, name=None)):
            vals = []
            for x in row:
                if pd.isna(x) or x is None:
                    vals.append("")
                elif isinstance(x, float):
                    vals.append(f"{x:.1f}")
                else:
                    vals.append(x)

            fatigue = str(vals[-1]) if vals else ""

            if fatigue in CLR["fatigue_colors"]:
                safe_tag = "fl_" + fatigue.replace(" ", "_")
                tags = (safe_tag,)
            else:
                tags = ("even" if count % 2 == 0 else "odd",)

            iid = tree.insert("", "end", values=vals, tags=tags)
            if count == 0:
                self._last_iid = iid
            if sel_id and str(vals[0]) == sel_id:
                restore_iid = iid

        try:
            if restore_iid:
                tree.selection_set(restore_iid)
            elif highlight_first and self._last_iid:
                tree.item(self._last_iid, tags=("last_tested",))
                tree.see(self._last_iid)
        except Exception as e:
            pass

    # ─────────────────────────────────────────────────────────────────
    # ─────────────────────────────────────────────────────────────────

    def refresh_ports(self):
        ports_info = serial.tools.list_ports.comports()
        ports = [f"{p.device} - {p.description}" for p in ports_info]
        self.port_combo["values"] = ports
        if ports and not self.selected_port.get():
            best = ports[0]
            for p in ports_info:
                if "Bluetooth" not in p.description and any(kw in p.description for kw in ["CP210", "CH340", "CH341", "FTDI", "USB", "UART"]):
                    best = f"{p.device} - {p.description}"
                    break
            self.selected_port.set(best)

    def _force_reconnect(self, event=None):
        target = self.selected_port.get().split(" - ")[0]
        if target:
            self._connect_serial(target)

    def _connect_serial(self, port):
        try:
            if self.serial_conn and self.serial_conn.is_open:
                self.serial_conn.close()
            self.serial_conn = serial.Serial(port, CONFIG["baudrate"], timeout=1)
            self._set_conn(f"✓ {port}", CLR["success"])
            return True
        except Exception as e:
            self._set_conn(f"✗ {port}", CLR["danger"])
            logging.error(f"Serial failed {port}: {e}")
            return False

    def _set_conn(self, t, c):
        # Thread-safe: schedule on main thread
        self.root.after(0, lambda: (self.conn_status.set(t),
                                    self._conn_label.configure(fg=c)))

    def _serial_loop(self):
        """Baca data dari serial port dan masukkan ke queue."""
        last_ports_cache = []
        while self.listener_running:
            try:
                ports  = serial.tools.list_ports.comports()
                current_ports_cache = [p.device for p in ports]
                if current_ports_cache != last_ports_cache:
                    last_ports_cache = current_ports_cache
                    self.root.after(0, self.refresh_ports)
                    
                if self.serial_conn is None or not self.serial_conn.is_open:
                    if CONFIG.get("auto_connect"):
                        target = self.selected_port.get().split(" - ")[0]
                        valid  = [p.device for p in ports]
                        
                        preferred_p = None
                        for p in ports:
                            if "Bluetooth" not in p.description and any(kw in p.description for kw in ["CP210", "CH340", "CH341", "FTDI", "USB", "UART"]):
                                preferred_p = p
                                break
                                
                        if preferred_p and target != preferred_p.device:
                            target = preferred_p.device
                            self.root.after(0, lambda dev=preferred_p.device, desc=preferred_p.description: self.selected_port.set(f"{dev} - {desc}"))
                            
                        if target in valid:
                            self._connect_serial(target)
                    if self.serial_conn is None or not self.serial_conn.is_open:
                        self._set_conn("Mencari ESP32…", CLR["warning"])
                        time.sleep(3)
                        continue
                if self.serial_conn.in_waiting > 0:
                    line = (self.serial_conn
                            .readline()
                            .decode("utf-8", errors="ignore")
                            .strip())
                    if line.startswith("R,A,T,I") or line.startswith("DATA"):
                        self._serial_queue.put(line)
            except serial.SerialException:
                self._set_conn("Terputus — reconnecting…", CLR["danger"])
                self.serial_conn = None
                time.sleep(2)
            except Exception as e:
                logging.error(f"Serial loop: {e}")
                time.sleep(2)

    def _serial_processor(self):
        """Proses data dari queue — terpisah dari I/O thread."""
        current_trials = []
        while self.listener_running:
            try:
                line = self._serial_queue.get(timeout=1)
                self._process_line(line, current_trials)
            except queue.Empty:
                continue
            except Exception as e:
                logging.error(f"Serial processor: {e}")

    def _process_line(self, line, current_trials):
        if not hasattr(self, 'active_test_window') or self.active_test_window is None:
            return

        parts = line.split(",")
        if len(parts) < 8:
            return
            
        # Kadang hardware kirim format yg berbeda, kita longgarkan sedikit pengecekannya
        if "TIME" not in parts and "FINAL" not in parts:
            return

        trial_or_final = parts[6]
        now_str        = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        user           = self.current_user.get()
        nik            = self.current_nik.get().strip().upper()
        usia           = self.current_usia.get()
        jabatan        = self.current_jabatan.get()
        divisi         = self.current_divisi.get()

        if trial_or_final == "FINAL":
            t_vals = [t["rt"] for t in current_trials if t["trial"] > 2]
            while len(t_vals) < 6:
                t_vals.append(0.0)
            
            t_vals = t_vals[:6]

            avg = round(float(np.mean(t_vals)), 1) if t_vals else 0.0
            fl  = classify_fatigue(avg)

            gender = self.current_gender.get()
            
            # Capture specific stimulus mode from hardware (e.g., Light, Sound, Random)
            hardware_mode = "ALAT"
            if len(parts) > 5 and parts[5].strip():
                hardware_mode = f"ALAT ({parts[5].strip()})"

            last_id = self._save_db(now_str, user, nik, usia, gender, jabatan, "", "", divisi,
                          hardware_mode,
                          t_vals[0], t_vals[1], t_vals[2],
                          t_vals[3], t_vals[4], t_vals[5],
                          avg, "OK", fl)
            if last_id is not None:
                self._save_excel(last_id, now_str, user, nik, usia, gender, jabatan, "", "", divisi,
                                 hardware_mode,
                                 t_vals[0], t_vals[1], t_vals[2],
                                 t_vals[3], t_vals[4], t_vals[5],
                                 avg, "OK", fl)
            current_trials.clear()
            self.root.after(0, lambda: self.refresh_dashboard(highlight=True))
            
            if self.active_test_window:
                user_name = self.current_user.get()
                self.root.after(0, lambda a=avg, f=fl, n=user_name:
                    self.active_test_window.show_result(n, a, f) 
                    if self.active_test_window else None)

            def _auto_medis():
                if self.active_test_window is None: return
                try:
                    # Simpan id untuk dipakai tombol Input Medis
                    self._pending_medis_id = last_id
                    # Tidak langsung buka medis — user klik tombol sendiri di window hasil
                except Exception as e:
                    logging.error(f"Auto medis failed: {e}")
                    self.root.after(0, lambda: self._update_status("⚠ Gagal buka form medis otomatis."))
            
            self.root.after(1000, _auto_medis)


        else:
            try:
                rt    = float(parts[7])
                t_idx = int(trial_or_final)
                current_trials.append({"trial": t_idx, "rt": rt})
                if self.active_test_window:
                    self.root.after(0, lambda i=t_idx: self.active_test_window.update_dot(i))
                    tag = " [PERCOBAAN]" if t_idx <= 2 else " [UJI ASLI]"
                    self.root.after(0, lambda r=rt, i=t_idx, tg=tag:
                                    self.active_test_window.set_status(
                                        f"⏳ Trial {i}/8 — {r} ms{tg}",
                                        CLR["warning"]))
            except Exception as e:
                logging.error(f"Parse: {e}")

    # ─────────────────────────────────────────────────────────────────
    # DB / EXCEL SAVE
    # ─────────────────────────────────────────────────────────────────
    def _save_db(self, tanggal, nama, nik, usia, jenis_kelamin, jabatan, contractor, plan,
                 divisi, mode, t1, t2, t3, t4, t5, t6, avg, status, fatigue):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                c.execute("""
                    INSERT INTO results
                      (tanggal,nama,nik,usia,jenis_kelamin,jabatan,contractor,plan,divisi,
                       mode,trial,t1,t2,t3,t4,t5,t6,avg_reaction,status,fatigue_level)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                             (tanggal, nama, nik, usia, jenis_kelamin, jabatan,
                              contractor, plan, divisi, mode, mode,
                              t1, t2, t3, t4, t5, t6, avg, status, fatigue))
                last_id = c.lastrowid
                conn.commit()
                logging.info(f"DB saved: id={last_id}, nik={nik}, fatigue={fatigue}, avg={avg:.1f}ms")
                return last_id
        except Exception as e:
            logging.error(f"DB save: {e}")
            self.root.after(0, lambda: self._update_status(f"⚠ Gagal simpan DB: {e}"))
            return None

    def _save_excel(self, db_id, tanggal, nama, nik, usia, jenis_kelamin, jabatan, contractor, plan,
                    divisi, mode, t1, t2, t3, t4, t5, t6, avg, status, fatigue):
        if not db_id: return
        def task():
            try:
                if not os.path.exists(EXCEL_PATH):
                    init_excel()
                wb = load_workbook(EXCEL_PATH)
                ws = wb.active
                total = t1 + t2 + t3 + t4 + t5 + t6

                ws.append([db_id, tanggal, nama, nik, usia, jenis_kelamin, jabatan,
                           t1, t2, t3, t4, t5, t6,
                           "", "", "", total, avg, "",
                           "", "", "", fatigue, "", "", "", ""])

                fill_map = {
                    "Normal":         ("2ECC71", "FFFFFF"),
                    "Fatigue Ringan": ("F1C40F", "000000"),
                    "Fatigue Sedang": ("E67E22", "FFFFFF"),
                    "Fatigue Berat":  ("E74C3C", "FFFFFF"),
                }
                if fatigue in fill_map:
                    fc, ft = fill_map[fatigue]
                    fill   = PatternFill(start_color=fc, end_color=fc,
                                         fill_type="solid")
                    font   = Font(color=ft)
                    for col in range(1, ws.max_column + 1):
                        ws.cell(ws.max_row, col).fill = fill
                        ws.cell(ws.max_row, col).font = font
                for col in range(10, 18):
                    c = ws.cell(ws.max_row, col)
                    if isinstance(c.value, (int, float)):
                        c.number_format = "0.0"
                wb.save(EXCEL_PATH)
            except Exception as e:
                logging.error(f"Excel save: {e}")
                self.root.after(0, lambda: self._update_status(f"⚠ Gagal simpan Excel: {e}"))
        threading.Thread(target=task, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────
    # DASHBOARD REFRESH
    # ─────────────────────────────────────────────────────────────────
    def refresh_dashboard(self, highlight=False):
        self._update_nik_list()
        self._update_stats()
        self._update_charts()
        self.update_table(highlight=highlight)
        self.update_table_medis()

    def _auto_refresh_loop(self):
        try:
            self.update_table(preserve=True)
            self.update_table_medis(preserve=True)
        except Exception:
            pass
        self.root.after(3000, self._auto_refresh_loop)

    def _update_stats(self):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                df = pd.read_sql_query(
                    "SELECT fatigue_level FROM results", conn)
            n      = len(df)
            normal = (df["fatigue_level"] == "Normal").sum()        if n else 0
            ringan = (df["fatigue_level"] == "Fatigue Ringan").sum() if n else 0
            sedang = (df["fatigue_level"] == "Fatigue Sedang").sum() if n else 0
            berat  = (df["fatigue_level"] == "Fatigue Berat").sum()  if n else 0

            self._kpi_total.config(text=str(n))
            self._kpi_normal.config(text=str(normal))
            self._kpi_ringan.config(text=str(ringan))
            self._kpi_sedang.config(
                text=str(sedang),
                fg=CLR["warning"] if sedang else CLR["accent"])
            self._kpi_berat.config(
                text=str(berat),
                fg=CLR["danger"] if berat else CLR["accent"])
        except Exception as e:
            logging.error(f"Stats: {e}")

    def refresh_charts(self):
        """Alias untuk _update_charts — dipanggil dari tab change."""
        self._update_charts()

    def refresh_charts(self):
        """Public alias untuk _update_charts."""
        self._update_charts()

    def _update_charts(self):
        try:
            limit = CONFIG.get("chart_limit", 500)
            with sqlite3.connect(DB_PATH) as conn:
                df = pd.read_sql_query(
                    f"SELECT tanggal, fatigue_level, divisi, jabatan, nama,"
                    f" (IFNULL(t1,0)+IFNULL(t2,0)+IFNULL(t3,0)"
                    f"  +IFNULL(t4,0)+IFNULL(t5,0)+IFNULL(t6,0)) AS total"
                    f" FROM results ORDER BY id DESC LIMIT {limit}", conn)
            for ax in [self.ax_daily, self.ax_pie, self.ax_divisi]:
                ax.clear()
                self._sax(ax)

            if not df.empty:
                # Weekly bar chart
                df["date"] = pd.to_datetime(df["tanggal"])
                df["week"] = df["date"].dt.strftime("%Y-W%V")
                weekly     = df["week"].value_counts().sort_index()
                self.ax_daily.bar(
                    weekly.index.astype(str), weekly.values,
                    color=CLR["accent"], width=.6, alpha=.85)
                self.ax_daily.set_title("Pengujian Mingguan", fontsize=9,
                                        color=CLR["text_main"])
                self.ax_daily.tick_params(axis="x", rotation=30, labelsize=6)
                self.ax_daily.set_ylabel("Jumlah", fontsize=7,
                                         color=CLR["text_muted"])

                # Pie — per fatigue_level
                if "fatigue_level" in df.columns:
                    df["fatigue_norm"] = (df["fatigue_level"].astype(str).str.strip())
                    df["fatigue_norm"] = df["fatigue_norm"].replace(
                        ["nan", "None", ""], "Lainnya")
                    fc = df["fatigue_norm"].value_counts().head(5)
                    
                    fatigue_pie_colors = {
                        "Normal": "#FFFFFF",         # Putih
                        "Fatigue Ringan": "#2ECC71", # Hijau
                        "Fatigue Sedang": "#F1C40F", # Kuning
                        "Fatigue Berat": "#E74C3C",  # Merah
                    }
                    pie_colors = [fatigue_pie_colors.get(label, "#BDC3C7") for label in fc.index]
                    
                    self.ax_pie.pie(
                        fc.values, labels=fc.index, autopct="%1.0f%%",
                        colors=pie_colors,
                        textprops={"fontsize": 7, "color": CLR["text_main"]},
                        wedgeprops={"linewidth": 1, "edgecolor": "#BDC3C7"})
                    self.ax_pie.set_title("Proporsi Tingkat Fatigue", fontsize=9,
                                          color=CLR["text_main"])

                # Total Responden per Divisi
                if "divisi" in df.columns:
                    df["divisi_norm"] = (df["divisi"].astype(str)
                                         .str.strip().str.upper())
                    df["divisi_norm"] = df["divisi_norm"].replace(
                        ["NAN", "NONE", ""], "TIDAK DIKETAHUI")
                    
                    dv = (df.groupby("divisi_norm")["nama"]
                          .nunique()
                          .sort_values(ascending=True)
                          .tail(10))
                    
                    self.ax_divisi.barh(
                        dv.index.astype(str), dv.values,
                        color=CLR["accent2"], alpha=.85, edgecolor="#BDC3C7")
                    self.ax_divisi.set_title(
                        "Total Responden per Divisi",
                        fontsize=9, color=CLR["text_main"])
                    self.ax_divisi.tick_params(axis="y", labelsize=7)
                    
                    for i, v in enumerate(dv.values):
                        self.ax_divisi.text(v, i, f" {v}", color=CLR["text_main"],
                                            va='center', fontsize=7, fontweight='bold')

            for fig in [self.fig_daily, self.fig_pie, self.fig_divisi]:
                fig.tight_layout()
                fig.canvas.draw()
        except Exception as e:
            logging.error(f"Charts: {e}")

    def update_table(self, reset=False, highlight=False, preserve=False):
        if reset:
            self.search_var.set("")
        q     = self.search_var.get().strip()
        limit = CONFIG.get("table_limit", 500)
        try:
            with sqlite3.connect(DB_PATH) as conn:
                if q:
                    df = pd.read_sql_query(f"""
                        SELECT id, tanggal, nama, nik, usia, jenis_kelamin, jabatan,
                               IFNULL(mode, '-') AS mode,
                               IFNULL(t1,0) AS t1, IFNULL(t2,0) AS t2,
                               IFNULL(t3,0) AS t3, IFNULL(t4,0) AS t4,
                               IFNULL(t5,0) AS t5, IFNULL(t6,0) AS t6,
                               IFNULL(keputusan, '-') AS status,
                               IFNULL(fatigue_level, '-') AS fatigue_level
                        FROM results
                        WHERE nama LIKE ? OR nik LIKE ?
                        ORDER BY id DESC LIMIT {limit}""",
                        conn, params=(f"%{q}%", f"%{q}%"))
                else:
                    df = pd.read_sql_query(f"""
                        SELECT id, tanggal, nama, nik, usia, jenis_kelamin, jabatan,
                               IFNULL(mode, '-') AS mode,
                               IFNULL(t1,0) AS t1, IFNULL(t2,0) AS t2,
                               IFNULL(t3,0) AS t3, IFNULL(t4,0) AS t4,
                               IFNULL(t5,0) AS t5, IFNULL(t6,0) AS t6,
                               IFNULL(keputusan, '-') AS status,
                               IFNULL(fatigue_level, '-') AS fatigue_level
                        FROM results
                        ORDER BY id DESC LIMIT {limit}""", conn)
            try:
                df["usia"] = pd.to_numeric(df["usia"], errors='coerce').astype('Int64')
            except Exception:
                pass
            self._populate_tree(self.tree, df, highlight_first=highlight, preserve_selection=preserve)
        except Exception as e:
            logging.error(f"Table update: {e}")

    def update_table_medis(self, reset=False, preserve=False):
        if reset:
            self.search_medis_var.set("")
        q = self.search_medis_var.get().strip()
        try:
            conn = sqlite3.connect(DB_PATH)
            sql = """
                SELECT id, tanggal, nama, nik, jenis_kelamin, jabatan,
                       IFNULL(avg_reaction, '') AS rt,
                       IFNULL(fatigue_level, '') AS fatigue,
                       IFNULL(td_sistol,  '') AS sistol,
                       IFNULL(td_diastol, '') AS diastol,
                       IFNULL(nadi,       '') AS nadi,
                       IFNULL(alcohol_test, '') AS alkohol,
                       IFNULL(kesimpulan_sistol,  '') AS kesp_sistol,
                       IFNULL(kesimpulan_diastol, '') AS kesp_diastol,
                       IFNULL(kesimpulan_nadi,    '') AS kesp_nadi,
                       IFNULL(keputusan,   '') AS keputusan,
                       IFNULL(rekomendasi, '') AS rekomendasi,
                       IFNULL(keterangan,  '') AS keterangan
                FROM results
                {where}
                ORDER BY id DESC LIMIT 500"""
            if q:
                df = pd.read_sql_query(
                    sql.format(where="WHERE nama LIKE ? OR nik LIKE ?"),
                    conn, params=(f"%{q}%", f"%{q}%"))
            else:
                df = pd.read_sql_query(sql.format(where=""), conn)
            conn.close()
            self._populate_tree(self.tree_medis, df, highlight_first=False, preserve_selection=preserve)
        except Exception as e:
            logging.error(f"Medis Table update: {e}")

    # ─────────────────────────────────────────────────────────────────
    # MEDIS DIALOG HELPERS
    # ─────────────────────────────────────────────────────────────────
    def _open_medis_dialog(self, from_medis_tab=False):
        tree = self.tree_medis if from_medis_tab else self.tree
        sel  = tree.selection()
        if not sel:
            messagebox.showwarning(
                "Pilih Data",
                "Pilih baris tes yang ingin ditambahkan data medisnya.")
            return
        v     = tree.item(sel[0], "values")
        db_id = v[0]   # column 0 = ID in both trees
        MedisInputDialog(self.root, db_id,
                         lambda: [self.update_table(),
                                  self.update_table_medis()],
                         app_mode=self.app_mode_var.get())

    def _on_medis_click(self, event):
        sel = self.tree_medis.selection()
        if not sel:
            return
        v = self.tree_medis.item(sel[0], "values")
        if len(v) > 0:
            self.ket_medis_id = v[0]
            self.ket_medis_var.set(v[-1])

    def _save_keterangan_cepat(self, event):
        if not self.ket_medis_id:
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("UPDATE results SET keterangan=? WHERE id=?",
                             (self.ket_medis_var.get(), self.ket_medis_id))
                conn.commit()
            self.update_table_medis()
            self._update_status("✅ Keterangan medis berhasil diperbarui.")
            self.tree_medis.focus_set()
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan keterangan: {e}")

    def _confirm_typing_delete(self, title, warning_msg):
        """Helper untuk konfirmasi hapus menggunakan PIN"""
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry("420x280")
        dlg.configure(bg=CLR["bg_panel"])
        dlg.grab_set()
        dlg.lift()
        dlg.focus_set()

        tk.Label(dlg, text=warning_msg,
                 font=("Segoe UI", 10), fg=CLR["text_main"], bg=CLR["bg_panel"],
                 justify="center").pack(pady=(16, 10))

        tk.Label(dlg, text="Masukkan PIN Konfirmasi:",
                 font=("Segoe UI", 10, "bold"),
                 fg=CLR["danger"], bg=CLR["bg_panel"]).pack(pady=(10, 6))
        
        var_confirm = tk.StringVar()
        ent = tk.Entry(dlg, textvariable=var_confirm,
                       font=("Segoe UI", 14, "bold"),
                       justify="center", width=14,
                       bg=CLR["bg_dark"], fg=CLR["text_main"],
                       insertbackground=CLR["text_main"], show="*")
        ent.pack(pady=4)
        ent.focus_set()

        confirmed = [False]

        def _do_delete():
            if _verify_pin(var_confirm.get().strip()):
                confirmed[0] = True
                dlg.destroy()
            else:
                ent.configure(bg="#5C1A1A")
                messagebox.showerror("Akses Ditolak", "PIN yang dimasukkan salah!", parent=dlg)
                var_confirm.set("")
                ent.configure(bg=CLR["bg_dark"])

        tk.Frame(dlg, bg=CLR["bg_panel"]).pack(expand=True)
        bf = tk.Frame(dlg, bg=CLR["bg_panel"])
        bf.pack(pady=10)
        HoverButton(bf, text="Konfirmasi",
                    font=("Segoe UI", 10, "bold"),
                    fg="white", bg=CLR["danger"],
                    relief="flat", cursor="hand2",
                    padx=14, pady=6, command=_do_delete).pack(side="left", padx=6)
        HoverButton(bf, text="Batal",
                    font=("Segoe UI", 10),
                    fg=CLR["text_muted"], bg=CLR["bg_panel"],
                    relief="flat", cursor="hand2",
                    padx=14, pady=6, command=dlg.destroy).pack(side="left", padx=6)

        ent.bind("<Return>", lambda e: _do_delete())
        dlg.wait_window()
        return confirmed[0]

    # ─────────────────────────────────────────────────────────────────
    # RESET DATA  (single definition — removed duplicate)
    # ─────────────────────────────────────────────────────────────────
    def _reset_data(self):
        msg = (
            "⚠️ PERHATIAN!\n\n"
            "Tindakan ini akan MENGHAPUS PERMANEN\n"
            "seluruh data hasil tes dari database.\n\n"
            "Excel lama akan di-backup otomatis.\n\n"
            "Lanjutkan ke konfirmasi akhir?"
        )
        if not self._confirm_typing_delete("Konfirmasi Hapus Semua Data", msg):
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("DELETE FROM results")
                try:
                    conn.execute(
                        "DELETE FROM sqlite_sequence WHERE name='results'")
                except Exception:
                    pass
                conn.commit()
            bk = ""
            if os.path.exists(EXCEL_PATH):
                bk = EXCEL_PATH.replace(
                    ".xlsx", f"_backup_{int(time.time())}.xlsx")
                os.rename(EXCEL_PATH, bk)
            self.refresh_dashboard()
            msg = "Semua data dihapus."
            if bk:
                msg += f"\nBackup: {os.path.basename(bk)}"
            messagebox.showinfo("Sukses", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ─────────────────────────────────────────────────────────────────
    # FILTER TAB LOGIC
    # ─────────────────────────────────────────────────────────────────
    def _get_filtered_df(self):
        clauses, params = [], []
        if self.filter_from.get().strip():
            clauses.append("DATE(tanggal)>=?")
            params.append(self.filter_from.get().strip())
        if self.filter_to.get().strip():
            clauses.append("DATE(tanggal)<=?")
            params.append(self.filter_to.get().strip())
        if self.filter_jabatan.get().strip():
            clauses.append("jabatan LIKE ?")
            params.append(f"%{self.filter_jabatan.get().strip()}%")
        if self.filter_fatigue.get() != "Semua":
            clauses.append("fatigue_level=?")
            params.append(self.filter_fatigue.get())
        if self.filter_keputusan.get() != "Semua":
            clauses.append("keputusan=?")
            params.append(self.filter_keputusan.get())

        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        conn  = sqlite3.connect(DB_PATH)
        df    = pd.read_sql_query(
            f"""SELECT id, tanggal, nama, nik, usia, jenis_kelamin, jabatan,
                       IFNULL(divisi, '') AS divisi, IFNULL(contractor, '') AS contractor, IFNULL(plan, '') AS plan,
                       IFNULL(td_sistol,  '') AS sistol,
                       IFNULL(td_diastol, '') AS diastol,
                       IFNULL(nadi,       '') AS nadi,
                       IFNULL(alcohol_test, '') AS alcohol_test,
                       IFNULL(t1,0) AS t1, IFNULL(t2,0) AS t2,
                       IFNULL(t3,0) AS t3, IFNULL(t4,0) AS t4,
                       IFNULL(t5,0) AS t5, IFNULL(t6,0) AS t6,
                       (IFNULL(t1,0)+IFNULL(t2,0)+IFNULL(t3,0)
                        +IFNULL(t4,0)+IFNULL(t5,0)+IFNULL(t6,0)) AS total,
                       avg_reaction,
                       IFNULL(fatigue_level,       '') AS fatigue_level,
                       IFNULL(kesimpulan_sistol,   '') AS kesp_sistol,
                       IFNULL(kesimpulan_diastol,  '') AS kesp_diastol,
                       IFNULL(kesimpulan_nadi,     '') AS kesp_nadi,
                       IFNULL(keputusan,           '') AS keputusan,
                       IFNULL(rekomendasi,         '') AS rekomendasi,
                       IFNULL(keterangan,          '') AS keterangan,
                       IFNULL(kesimpulan_alkohol,  '') AS kesimpulan_alkohol
                FROM results {where}
                ORDER BY id DESC""",
            conn, params=params)
        conn.close()
        
        try:
            # Convert usia to nullable integer to remove .0 decimals
            df["usia"] = pd.to_numeric(df["usia"], errors='coerce').astype('Int64')
        except Exception:
            pass
            
        return df

    def _open_medis_dialog_by_id(self, db_id):
        MedisInputDialog(self.root, db_id, lambda: [self.update_table(), self.update_table_medis(), self._apply_filter()], app_mode=self.app_mode_var.get())

    def _on_tab_changed(self, event=None):
        try:
            idx = self.notebook.index(self.notebook.select())
            if idx == 0:
                self.update_table()
            elif idx == 1:
                self.update_table_medis()
            elif idx == 2:
                self.refresh_charts()
            elif idx == 3:
                self._apply_filter()
            elif idx == 4:
                self._mk_refresh(is_kontraktor=False)
            elif idx == 5:
                self._mk_refresh(is_kontraktor=True)
            # idx 6 = Settings, idx 7 = Tentang — tidak perlu refresh
        except Exception:
            pass

    def _apply_filter(self):
        try:
            df = self._get_filtered_df()
            # FIX: display_cols must match actual column names returned by query
            display_cols = ["id", "tanggal", "nama", "nik", "jenis_kelamin", "jabatan",
                            "t1", "t2", "t3", "t4", "t5", "t6", "avg_reaction",
                            "total", "sistol", "diastol", "nadi",
                            "fatigue_level", "keputusan"]
            display_df = df[display_cols].head(500)
            self._populate_tree(self.tree_filter, display_df)
            self._update_status(
                f"Filter — {len(df)} baris (max 500 ditampilkan).")
        except Exception as e:
            messagebox.showerror("Error Filter", str(e))

    def _reset_filter(self):
        self.filter_from.set("")
        self.filter_to.set("")
        self.filter_jabatan.set("")
        self.filter_fatigue.set("Semua")
        self.filter_keputusan.set("Semua")
        for item in self.tree_filter.get_children():
            self.tree_filter.delete(item)
        self._update_status("Filter direset.")

    def _search(self):
        self.update_table()

    def _delete_selected(self):
        sel = self.tree_filter.selection()
        if not sel:
            messagebox.showwarning("Pilih Data",
                                   "Pilih baris yang ingin dihapus.")
            return

        # Ambil ID database dari GUI SEBELUM dialog muncul (menghindari error background refresh)
        to_delete_ids = []
        for item in sel:
            v = self.tree_filter.item(item, "values")
            to_delete_ids.append(v[0])

        pin = simpledialog.askstring("Admin PIN", "Masukkan PIN Admin:", show="*", parent=self.root)
        if not _verify_pin(pin or ""):
            messagebox.showerror("Akses Ditolak", "PIN Admin salah!", parent=self.root)
            return

        if not messagebox.askyesno("Konfirmasi", f"Hapus {len(sel)} data terpilih?"):
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                for db_id in to_delete_ids:
                    c.execute("DELETE FROM results WHERE id=?", (db_id,))
                conn.commit()
            self._apply_filter()
            self.refresh_dashboard()
            messagebox.showinfo("Berhasil", f"{len(sel)} data dihapus.")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    # ─────────────────────────────────────────────────────────────────
    # DOUBLE-CLICK → History popup
    # ─────────────────────────────────────────────────────────────────
    def _on_tree_double_click(self, event):
        self._open_history_from_tree(self.tree)

    def _on_filter_tree_double_click(self, event):
        self._open_history_from_tree(self.tree_filter)

    def _open_history_from_tree(self, tree):
        sel = tree.selection()
        if not sel:
            return
        v = tree.item(sel[0], "values")
        # FIX: col 0=ID, col 2=Nama, col 3=NIK  (same for both trees)
        nik  = str(v[3]) if len(v) > 3 else ""
        nama = str(v[2]) if len(v) > 2 else ""
        if nik:
            HistoryPopup(self.root, nik, nama)

    # ─────────────────────────────────────────────────────────────────
    # MASTER KARYAWAN / KONTRAKTOR CRUD
    # ─────────────────────────────────────────────────────────────────
    def _mk_refresh(self, is_kontraktor=False):
        tree       = self.tree_master_kontraktor if is_kontraktor else self.tree_master
        search_var = self._mkon_search           if is_kontraktor else self._mk_search
        table      = "master_kontraktor"         if is_kontraktor else "master_karyawan"

        for item in tree.get_children():
            tree.delete(item)
        q = search_var.get().strip()
        try:
            with sqlite3.connect(DB_PATH) as conn:
                if q:
                    df = pd.read_sql_query(
                        f"SELECT id, nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan"
                        f" FROM {table}"
                        f" WHERE nama LIKE ? OR nik LIKE ?"
                        f" ORDER BY id DESC",
                        conn, params=(f"%{q}%", f"%{q}%"))
                else:
                    df = pd.read_sql_query(
                        f"SELECT id, nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan"
                        f" FROM {table} ORDER BY id DESC LIMIT 300",
                        conn)
            for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
                vals = [str(x) if x is not None else "" for x in row]
                vals[0] = str(i)
                # Note: df in _mk_refresh has 'id' at index 0 because of query 'SELECT id, ...'
                # row[0] is the 'id'
                tree.insert("", "end", iid=str(row[0]), values=vals,
                            tags=("even" if i % 2 == 0 else "odd",))
        except Exception as e:
            logging.error(f"Master refresh: {e}")

    def _mk_add(self, is_kontraktor=False):
        title = "Tambah Kontraktor" if is_kontraktor else "Tambah Karyawan"
        self._mk_form(title, {}, is_kontraktor)

    def _mk_edit(self, is_kontraktor=False):
        tree  = self.tree_master_kontraktor if is_kontraktor else self.tree_master
        title = "Edit Kontraktor"           if is_kontraktor else "Edit Karyawan"
        sel   = tree.selection()
        if not sel:
            messagebox.showwarning("Pilih", "Pilih baris terlebih dahulu.")
            return
        v = tree.item(sel[0], "values")
        self._mk_form(title, {
            "id":            sel[0],
            "nama":          v[1],
            "nik":           v[2],
            "usia":          v[3],
            "jenis_kelamin": v[4],
            "jabatan":       v[5],
            "info":          v[6],
        }, is_kontraktor)

    def _mk_form(self, title, data, is_kontraktor=False):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("380x360")
        win.configure(bg=CLR["bg_panel"])
        win.grab_set()

        tk.Label(win, text=title, font=("Segoe UI", 12, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(14, 8))
        form  = tk.Frame(win, bg=CLR["bg_panel"])
        form.pack(padx=20, fill="x")
        vcmd  = (win.register(lambda v: v == "" or v.isdigit()), "%P")
        vars_ = {}

        info_label = "Nama PT / Perusahaan" if is_kontraktor else "Plan/Devisi/Perusahaan"
        fields = [
            ("Nama",          "nama",          False),
            ("NIK",           "nik",           False),
            ("Usia",          "usia",          True),
            ("Jenis Kelamin", "jenis_kelamin", False),
            ("Jabatan",       "jabatan",       False),
            (info_label,      "info",          False),
        ]
        for i, (lbl, key, num) in enumerate(fields):
            tk.Label(form, text=lbl, font=("Segoe UI", 9),
                     fg=CLR["text_muted"], bg=CLR["bg_panel"],
                     width=14, anchor="w").grid(row=i, column=0, pady=3)
            v = tk.StringVar(value=data.get(key, ""))
            
            if key == "jenis_kelamin":
                e = ttk.Combobox(form, textvariable=v, values=["Laki-laki", "Perempuan"],
                                 font=("Segoe UI", 9), width=23)
            else:
                e = tk.Entry(form, textvariable=v, font=("Segoe UI", 9),
                             bg=CLR["bg_card"], fg=CLR["text_main"],
                             relief="flat", bd=4, width=24)
                if num:
                    e.configure(validate="key", validatecommand=vcmd)
            
            e.grid(row=i, column=1, pady=3, padx=4)
            vars_[key] = v

        def _save():
            vars_["nik"].set(vars_["nik"].get().strip().upper())
            table = "master_kontraktor" if is_kontraktor else "master_karyawan"
            if table not in ("master_karyawan", "master_kontraktor"):
                messagebox.showerror("Error", "Invalid table name")
                return
            try:
                with sqlite3.connect(DB_PATH) as conn:
                    if data.get("id"):
                        old_nik_row = conn.execute(
                            f"SELECT nik FROM {table} WHERE id=?",
                            (data["id"],)).fetchone()
                        old_nik = old_nik_row[0] if old_nik_row else None
                        if old_nik and old_nik != vars_["nik"].get():
                            conn.execute(
                                "UPDATE results SET nik=?, nama=? WHERE nik=?",
                                (vars_["nik"].get(), vars_["nama"].get(), old_nik))
                        conn.execute(
                            f"UPDATE {table}"
                            " SET nama=?, nik=?, usia=?, jenis_kelamin=?, jabatan=?, info_pekerjaan=?"
                            " WHERE id=?",
                            (vars_["nama"].get(), vars_["nik"].get(),
                             vars_["usia"].get(), vars_["jenis_kelamin"].get(), vars_["jabatan"].get(),
                             vars_["info"].get(), data["id"]))
                    else:
                        conn.execute(
                            f"INSERT INTO {table}"
                            " (nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan)"
                            " VALUES (?,?,?,?,?,?)",
                            (vars_["nama"].get(), vars_["nik"].get(),
                             vars_["usia"].get(), vars_["jenis_kelamin"].get(), vars_["jabatan"].get(),
                             vars_["info"].get()))
                    conn.commit()
                self._sync_master_excel()
                self._mk_refresh(is_kontraktor)
                self._update_nik_list()
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        bf = tk.Frame(win, bg=CLR["bg_panel"])
        bf.pack(pady=10)
        HoverButton(bf, text="Simpan", command=_save,
                    font=("Segoe UI", 10, "bold"),
                    bg=CLR["accent"], fg="white",
                    relief="flat", padx=16, pady=4,
                    cursor="hand2").pack(side="left", padx=6)
        HoverButton(bf, text="Batal", command=win.destroy,
                    font=("Segoe UI", 10),
                    bg=CLR["bg_card"], fg=CLR["text_muted"],
                    relief="flat", padx=16, pady=4,
                    cursor="hand2").pack(side="left", padx=6)

    def _mk_delete(self, is_kontraktor=False):
        tree  = self.tree_master_kontraktor if is_kontraktor else self.tree_master
        table = "master_kontraktor"         if is_kontraktor else "master_karyawan"
        label = "kontraktor"                if is_kontraktor else "karyawan"
        sel   = tree.selection()
        if not sel:
            messagebox.showwarning("Pilih", "Pilih baris.")
            return

        # Ambil ID database dari GUI SEBELUM dialog muncul
        to_delete_data = []
        for item in sel:
            v = tree.item(item, "values")
            to_delete_data.append((item, v[2])) # (db_id, NIK)

        pin = simpledialog.askstring("Admin PIN", "Masukkan PIN Admin:", show="*", parent=self.root)
        if not _verify_pin(pin or ""):
            messagebox.showerror("Akses Ditolak", "PIN Admin salah!", parent=self.root)
            return

        if not messagebox.askyesno(
                "Konfirmasi",
                f"Hapus {label} terpilih beserta SEMUA data tes mereka?"):
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                for db_id, nik in to_delete_data:
                    c.execute(f"DELETE FROM {table} WHERE id=?", (db_id,))
                    if nik:
                        c.execute("DELETE FROM results WHERE nik=?", (nik,))
                conn.commit()
            self._sync_master_excel()
            self._mk_refresh(is_kontraktor)
            self._update_nik_list()
            self.refresh_dashboard()
            messagebox.showinfo(
                "Berhasil",
                f"Data {label} dan riwayat tesnya berhasil dihapus.")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def _mk_delete_all(self, is_kontraktor=False):
        table = "master_kontraktor" if is_kontraktor else "master_karyawan"
        label = "Master Kontraktor" if is_kontraktor else "Master Karyawan"
        
        msg = (
            f"⚠️ PERHATIAN!\n\n"
            f"Anda akan menghapus SEMUA data di {label}\n"
            f"beserta seluruh riwayat tes yang terkait dengan mereka.\n\n"
            f"Tindakan ini tidak dapat dibatalkan.\n"
            f"Lanjutkan ke konfirmasi akhir?"
        )
        if not self._confirm_typing_delete(f"Reset {label}", msg):
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                c.execute(
                    f"DELETE FROM results"
                    f" WHERE nik IN (SELECT nik FROM {table})")
                c.execute(f"DELETE FROM {table}")
                conn.commit()
            self._sync_master_excel()
            self._mk_refresh(is_kontraktor)
            self._update_nik_list()
            self.refresh_dashboard()
            messagebox.showinfo(
                "Berhasil",
                f"Semua {label} dan riwayat tes terkait berhasil dihapus.")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    # Kontraktor wrappers
    def _mkon_refresh(self):    self._mk_refresh(is_kontraktor=True)
    def _mkon_add(self):        self._mk_add(is_kontraktor=True)
    def _mkon_edit(self):       self._mk_edit(is_kontraktor=True)
    def _mkon_delete(self):     self._mk_delete(is_kontraktor=True)
    def _mkon_delete_all(self): self._mk_delete_all(is_kontraktor=True)

    # ─────────────────────────────────────────────────────────────────
    # BARCODE / ID-CARD SCANNER (global key handler)
    # ─────────────────────────────────────────────────────────────────
    def _global_key_handler(self, event):
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R",
                             "Alt_L", "Alt_R", "Caps_Lock", "Tab"):
            return

        now = time.time()
        # Scanner input sangat cepat (< 0.05s antar karakter).
        # Hapus filter focus_get agar scanner bisa overwrite widget langsung,
        # tanpa harus manual block/Ctrl+A di NIK combobox.
        if now - self._last_key_time > 0.05:
            self._scan_buffer = ""

        self._last_key_time = now

        if event.keysym == "Return":
            if len(self._scan_buffer) >= 4:
                self.current_nik.set(self._scan_buffer)
                self._autofill()
            self._scan_buffer = ""
        elif event.char and event.char.isprintable():
            self._scan_buffer += event.char

    # ─────────────────────────────────────────────────────────────────
    # IMPORT MASTER
    # ─────────────────────────────────────────────────────────────────
    def _import_master(self, is_kontraktor=False):
        table = "master_kontraktor" if is_kontraktor else "master_karyawan"
        label = "Kontraktor"        if is_kontraktor else "Karyawan"
        path  = filedialog.askopenfilename(
            title=f"Pilih File Master {label}",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"),
                       ("Semua",     "*.*")])
        if not path:
            return
        try:
            df = (pd.read_csv(path)
                  if path.lower().endswith(".csv")
                  else pd.read_excel(path))
            df.columns = [
                str(c).lower().strip().replace(" ", "_") for c in df.columns]
            df = df.where(pd.notnull(df), None)
            with sqlite3.connect(DB_PATH) as conn:
                n    = 0
                for _, row in df.iterrows():
                    nama = row.get("nama")
                    if not nama or pd.isna(nama):
                        continue
                    nik  = str(row.get("nik",  "")).strip().upper()
                    if not nik or nik in ("NONE", "NAN", "NULL", "-", "NONE/NULL"):
                        continue
                    
                    tanggal_lahir = ""
                    for col in row.keys():
                        if "tanggal_lahir" in str(col) or "tanggal lahir" in str(col):
                            tanggal_lahir = row.get(col)
                            break
                    
                    if pd.notna(tanggal_lahir) and str(tanggal_lahir).strip():
                        tanggal_lahir = str(tanggal_lahir).strip()[:10]
                    else:
                        tanggal_lahir = ""

                    usia = row.get("usia", "")
                    if tanggal_lahir:
                        try:
                            birth_year = int(tanggal_lahir[:4])
                            current_year = datetime.date.today().year
                            usia = str(current_year - birth_year)
                        except Exception:
                            pass
                    
                    if pd.isna(usia): usia = ""
                    
                    jk   = (row.get("jenis_kelamin") or row.get("gender") or "")
                    jab  = (row.get("jabatan") or "")
                    info = (row.get("info_pekerjaan")
                            or row.get("perusahaan/bagian")
                            or row.get("keterangan")
                            or row.get("divisi")
                            or row.get("departemen")
                            or row.get("contractor")
                            or "")
                    try:
                        conn.execute(
                            f"INSERT INTO {table}"
                            " (nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan, tanggal_lahir) VALUES (?,?,?,?,?,?,?)"
                            " ON CONFLICT(nik) DO UPDATE SET"
                            " nama=excluded.nama, usia=excluded.usia, jenis_kelamin=excluded.jenis_kelamin,"
                            " jabatan=excluded.jabatan, info_pekerjaan=excluded.info_pekerjaan, tanggal_lahir=excluded.tanggal_lahir",
                            (str(nama), str(nik), str(usia), str(jk), str(jab), str(info), str(tanggal_lahir)))
                        n += 1
                    except Exception as e:
                        logging.error(f"Master import row: {e}")
                conn.commit()
            self._sync_master_excel()
            self._mk_refresh(is_kontraktor)
            self._update_nik_list()
            messagebox.showinfo(
                "Berhasil",
                f"{n} data {label.lower()} berhasil diimpor/diperbarui.")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))
    def _select_master_for_test(self, tree):
        sel = tree.selection()
        if not sel: return
        nik = str(tree.item(sel[0])["values"][2])
        if nik:
            self.current_nik.set(nik)
            self._autofill()
            self.notebook.select(0)

    def _export_master(self, is_kontraktor=False):
        table = "master_kontraktor" if is_kontraktor else "master_karyawan"
        label = "Kontraktor"        if is_kontraktor else "Karyawan"
        dest  = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Master_{label}.xlsx",
            title=f"Export Master {label}")
        if not dest:
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                df   = pd.read_sql_query(
                    f"SELECT nama as Nama, nik as NIK, usia as Usia,"
                    f" jenis_kelamin as 'Jenis Kelamin', jabatan as Jabatan,"
                    f" info_pekerjaan as 'Bagian/PT/Perusahaan' FROM {table}", conn)
            df.to_excel(dest, index=False)
            if messagebox.askyesno(
                    "Berhasil",
                    f"Data Master {label} berhasil diexport ke:\n"
                    f"{os.path.basename(dest)}\n\nBuka file sekarang?"):
                _open_file(dest)
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def _import_master_kontraktor(self): self._import_master(is_kontraktor=True)

    def _download_template_master(self, is_kontraktor=False):
        label = "Kontraktor" if is_kontraktor else "Karyawan"
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Template_Master_{label}.xlsx",
            title=f"Download Template Master {label}")
        if not dest:
            return
        try:
            df = pd.DataFrame({
                "Nama": [],
                "NIK": [],
                "Usia": [],
                "Jenis Kelamin": [],
                "Jabatan": [],
                "Perusahaan/Bagian": []
            })
            if is_kontraktor:
                df.loc[0] = ["Contoh Kontraktor 1", "KTR-001", "35", "Laki-laki", "Teknisi", "PT Maju Jaya"]
                df.loc[1] = ["Contoh Kontraktor 2", "KTR-002", "28", "Perempuan", "Pengawas", "PT Sentosa Abadi"]
            else:
                df.loc[0] = ["Contoh Karyawan 1", "KRY-001", "30", "Laki-laki", "Operator", "Plant A"]
                df.loc[1] = ["Contoh Karyawan 2", "KRY-002", "25", "Perempuan", "Staff Admin", "Plant B"]
            
            df.to_excel(dest, index=False)
            messagebox.showinfo("Berhasil", f"Template {label} berhasil disimpan di:\n{dest}")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def _download_template_master_karyawan(self): self._download_template_master(is_kontraktor=False)
    def _download_template_master_kontraktor(self): self._download_template_master(is_kontraktor=True)
    def _export_master_kontraktor(self): self._export_master(is_kontraktor=True)

    def _sync_master_excel(self):
        """Sync master_karyawan & master_kontraktor → background Excel."""
        def _task():
            try:
                conn = sqlite3.connect(DB_PATH)
                df1  = pd.read_sql_query(
                    "SELECT nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan"
                    " FROM master_karyawan", conn)
                df2  = pd.read_sql_query(
                    "SELECT nama, nik, usia, jenis_kelamin, jabatan, info_pekerjaan"
                    " FROM master_kontraktor", conn)
                conn.close()
                with pd.ExcelWriter(EXCEL_MK_PATH) as writer:
                    df1.to_excel(writer, sheet_name="Karyawan",   index=False)
                    df2.to_excel(writer, sheet_name="Kontraktor", index=False)
            except Exception as e:
                logging.error(f"Sync Master Excel failed: {e}")
        
        # Offload to background thread
        threading.Thread(target=_task, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────
    # NIK AUTOFILL
    # ─────────────────────────────────────────────────────────────────
    def _update_nik_list(self):
        try:
            with sqlite3.connect(DB_PATH) as conn:
                d2 = pd.read_sql_query(
                    "SELECT DISTINCT nik, nama FROM master_karyawan WHERE nik IS NOT NULL AND nik!=''", conn)
                d3 = pd.read_sql_query(
                    "SELECT DISTINCT nik, nama FROM master_kontraktor WHERE nik IS NOT NULL AND nik!=''", conn)
                d1 = pd.read_sql_query(
                    "SELECT DISTINCT nik, nama FROM results WHERE nik IS NOT NULL AND nik!=''", conn)
            
            seen = set()
            nik_options = []
            for df in [d2, d3, d1]:
                for _, row in df.iterrows():
                    nik = str(row['nik']).strip()
                    nama = str(row['nama']).strip() if ('nama' in row and pd.notna(row['nama'])) else ""
                    if nik and nik not in seen:
                        seen.add(nik)
                        label = f"{nik} - {nama}" if (nama and nama.lower() != 'none') else nik
                        nik_options.append(label)

            if hasattr(self, "nik_cb") and isinstance(self.nik_cb, ttk.Combobox):
                self.nik_cb["values"] = sorted(nik_options)
        except Exception as e:
            logging.error(f"Error loading NIK list: {e}")

    def _autofill(self, event=None):
        raw_val = self.current_nik.get().strip().upper()
        if not raw_val:
            return
        # Extract NIK if formatted as 'NIK - NAMA'
        nik = raw_val.split(" - ")[0].strip()
        if " - " in raw_val:
            self.current_nik.set(nik)
        try:
            conn = sqlite3.connect(DB_PATH)
            # Try master_karyawan first
            df = pd.read_sql_query(
                "SELECT nama, usia, tanggal_lahir, jenis_kelamin, jabatan, info_pekerjaan AS divisi"
                " FROM master_karyawan WHERE nik=? LIMIT 1",
                conn, params=(nik,))
            if df.empty:
                df = pd.read_sql_query(
                    "SELECT nama, usia, tanggal_lahir, jenis_kelamin, jabatan, info_pekerjaan AS divisi"
                    " FROM master_kontraktor WHERE nik=? LIMIT 1",
                    conn, params=(nik,))
            if df.empty:
                # Fallback: last result
                df = pd.read_sql_query(
                    "SELECT nama, usia, null as tanggal_lahir, jenis_kelamin, jabatan, divisi"
                    " FROM results WHERE nik=? ORDER BY id DESC LIMIT 1",
                    conn, params=(nik,))
            conn.close()
            if not df.empty:
                # Need to convert to dict or modify the Series to avoid SettingWithCopyWarning
                r = df.iloc[0].to_dict()
                if r.get("tanggal_lahir") and pd.notna(r["tanggal_lahir"]):
                    try:
                        birth_year = int(str(r["tanggal_lahir"]).strip()[:4])
                        r["usia"] = str(datetime.date.today().year - birth_year)
                    except Exception:
                        pass
                
                if r.get("nama") and pd.notna(r["nama"]):          self.current_user.set(str(r["nama"]))
                if r.get("usia") and pd.notna(r["usia"]):          self.current_usia.set(str(r["usia"]))
                if r.get("jenis_kelamin"): self.current_gender.set(str(r["jenis_kelamin"]))
                if r.get("jabatan"):       self.current_jabatan.set(str(r["jabatan"]))
                if r.get("divisi"):        self.current_divisi.set(str(r["divisi"]))
                self._update_status(
                    f"Auto-fill: {r['nama']} (NIK {nik})")
            else:
                self.current_user.set("")
                self.current_usia.set("")
                self.current_gender.set("")
                self.current_divisi.set("")
                self._update_status(
                    f"NIK {nik} tidak ditemukan, silakan isi manual.")
        except Exception as e:
            logging.error(f"Autofill: {e}")

    # ─────────────────────────────────────────────────────────────────
    # MANUAL INPUT
    # ─────────────────────────────────────────────────────────────────
    def _open_manual_input(self):
        ManualInputDialog(self.root, self._save_manual)

    def _save_manual(self, row):
        ts  = row["ts"]
        avg = row["avg"]
        fl  = row["fatigue"]
        now = row["tanggal"]
        last_id = self._save_db(
            now, row["nama"], row["nik"], row["usia"], row.get("jenis_kelamin", ""),
            row["jabatan"], "", "", row["divisi"],
            "MANUAL",
            ts[0], ts[1], ts[2], ts[3], ts[4], ts[5],
            avg, "OK", fl)
        if last_id is not None:
            self._save_excel(
                last_id,
                now, row["nama"], row["nik"], row["usia"], row.get("jenis_kelamin", ""),
                row["jabatan"], "", "", row["divisi"],
                "MANUAL",
                ts[0], ts[1], ts[2], ts[3], ts[4], ts[5],
                avg, "OK", fl)
        self.refresh_dashboard(highlight=True)
        self._update_status(f"Data manual disimpan: {row['nama']} — {fl}")

    # ─────────────────────────────────────────────────────────────────
    # SETTINGS
    # ─────────────────────────────────────────────────────────────────
    def _save_settings(self):
        name_changed = (self._cfg_app_name.get() != CONFIG.get("app_name", "FAST"))
        
        if name_changed:
            master = simpledialog.askstring("Otorisasi", "Masukkan Master Key untuk mengubah identitas aplikasi:", show="*", parent=self.root)
            if master != "ITPSHE#2026":
                messagebox.showerror("Akses Ditolak", "Master Key salah atau dibatalkan!")
                return
                
        try:
            baud = int(self._cfg_baud.get())
        except ValueError:
            messagebox.showerror("Error", "Baud rate harus angka.")
            return
        try:
            age_thr = int(self._cfg_age_thr.get())
            if not (1 <= age_thr <= 120):
                raise ValueError("Batas usia di luar range")
        except ValueError:
            messagebox.showerror("Error", "Batas usia harus angka antara 1–120.")
            return

        try:
            limit_test = int(self._cfg_limit_test.get())
            if not (1 <= limit_test <= 10):
                raise ValueError()
        except ValueError:
            messagebox.showerror("Error", "Batas tes per hari harus angka antara 1–10.")
            return

        CONFIG["baudrate"]              = baud
        CONFIG["auto_connect"]          = self._cfg_auto.get()
        CONFIG["auto_backup"]           = self._cfg_backup.get()
        CONFIG["app_name"]              = self._cfg_app_name.get()
        CONFIG["company_name"]          = self._cfg_company.get().strip()[:100]
        CONFIG["created_by"]            = self._cfg_created_by.get().strip()[:100]
        if "app_title" in CONFIG:
            del CONFIG["app_title"]
        CONFIG["age_warning_threshold"] = age_thr
        CONFIG["daily_test_limit"]      = limit_test
        _save_config(CONFIG)
            
        self.root.title(f"{CONFIG['app_name']}")
        if hasattr(self, 'sidebar_app_name_label'):
            self.sidebar_app_name_label.config(text=CONFIG['app_name'])
        if hasattr(self, 'about_app_name_label'):
            self.about_app_name_label.config(text=CONFIG['app_name'])
        if hasattr(self, 'about_created_by_label'):
            self.about_created_by_label.config(text=f"Dibuat Oleh: {CONFIG['created_by']}")
        if hasattr(self, 'about_app_title_label'):
            self.about_app_title_label.config(text="")
        self._update_status("✅ Pengaturan disimpan.")
        messagebox.showinfo("Berhasil", "Pengaturan berhasil disimpan.\nAplikasi mungkin perlu di-restart agar semua perubahan aktif.")

    def _change_pin_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Ubah PIN Admin")
        dlg.geometry("340x360")
        dlg.configure(bg=CLR["bg_panel"])
        dlg.grab_set()

        tk.Label(dlg, text="Ubah PIN Hapus Data", font=("Segoe UI", 12, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(16, 10))

        tk.Label(dlg, text="Masukkan PIN Lama:", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        old_var = tk.StringVar()
        tk.Entry(dlg, textvariable=old_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        tk.Label(dlg, text="Masukkan PIN Baru (Min. 4 angka):", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        new1_var = tk.StringVar()
        tk.Entry(dlg, textvariable=new1_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        tk.Label(dlg, text="Konfirmasi PIN Baru:", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        new2_var = tk.StringVar()
        tk.Entry(dlg, textvariable=new2_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        def save_pin():
            if not _verify_pin(old_var.get().strip()):
                messagebox.showerror("Akses Ditolak", "PIN Lama salah!", parent=dlg)
                return
            new1 = new1_var.get().strip()
            if len(new1) < 4:
                messagebox.showerror("Peringatan", "PIN Baru minimal 4 karakter!", parent=dlg)
                return
            if new1 != new2_var.get().strip():
                messagebox.showerror("Peringatan", "Konfirmasi PIN Baru tidak cocok!", parent=dlg)
                return
            
            CONFIG["admin_pin"] = _hash_pin(new1)
            CONFIG["admin_pin_hashed"] = True
            try:
                _save_config(CONFIG)
                messagebox.showinfo("Sukses", "PIN berhasil diubah!\n\nPastikan Anda mengingat PIN yang baru.", parent=dlg)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan PIN: {e}", parent=dlg)

        HoverButton(dlg, text="Simpan PIN Baru", command=save_pin,
                    font=("Segoe UI", 10, "bold"), bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2", padx=16, pady=6).pack(pady=20)

    def _reset_pin_dialog(self):
        master_key = simpledialog.askstring("Reset PIN", "Masukkan Master Key untuk mereset PIN:", show="*", parent=self.root)
        if master_key == "ITPSHE#2026": # example master key
            CONFIG["admin_pin"] = _hash_pin("1234")
            CONFIG["admin_pin_hashed"] = True
            try:
                _save_config(CONFIG)
                messagebox.showinfo("Berhasil", "PIN berhasil di-reset ke default: 1234", parent=self.root)
            except Exception as e:
                messagebox.showerror("Error", f"Gagal reset PIN: {e}", parent=self.root)
        elif master_key is not None:
            messagebox.showerror("Akses Ditolak", "Master Key salah!", parent=self.root)

    def _change_paramedis_pin_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Ubah PIN Mode Paramedis")
        dlg.geometry("340x360")
        dlg.configure(bg=CLR["bg_panel"])
        dlg.grab_set()

        tk.Label(dlg, text="Ubah PIN Akses Mode Paramedis", font=("Segoe UI", 12, "bold"),
                 fg=CLR["text_main"], bg=CLR["bg_panel"]).pack(pady=(16, 10))

        tk.Label(dlg, text="Masukkan PIN Lama:", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        old_var = tk.StringVar()
        tk.Entry(dlg, textvariable=old_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        tk.Label(dlg, text="Masukkan PIN Baru (Min. 4 karakter):", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        new1_var = tk.StringVar()
        tk.Entry(dlg, textvariable=new1_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        tk.Label(dlg, text="Konfirmasi PIN Baru:", bg=CLR["bg_panel"], fg=CLR["text_muted"], font=("Segoe UI", 10)).pack(pady=(10, 2))
        new2_var = tk.StringVar()
        tk.Entry(dlg, textvariable=new2_var, show="*", font=("Segoe UI", 11), width=24,
                 bg="#F1F5F9", fg=CLR["text_main"], relief="flat", bd=0,
                 highlightbackground="#CBD5E1", highlightcolor=CLR["accent"], highlightthickness=1).pack(pady=2)

        def save_pin():
            if not _verify_paramedis_pin(old_var.get().strip()):
                messagebox.showerror("Akses Ditolak", "PIN Lama salah!", parent=dlg)
                return
            new1 = new1_var.get().strip()
            if len(new1) < 4:
                messagebox.showerror("Peringatan", "PIN Baru minimal 4 karakter!", parent=dlg)
                return
            if new1 != new2_var.get().strip():
                messagebox.showerror("Peringatan", "Konfirmasi PIN Baru tidak cocok!", parent=dlg)
                return

            CONFIG["paramedis_pin"] = _hash_pin(new1)
            CONFIG["paramedis_pin_hashed"] = True
            try:
                _save_config(CONFIG)
                messagebox.showinfo("Sukses", "PIN Mode Paramedis berhasil diubah!\n\nPastikan Anda mengingat PIN yang baru.", parent=dlg)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Gagal menyimpan PIN: {e}", parent=dlg)

        HoverButton(dlg, text="Simpan PIN Baru", command=save_pin,
                    font=("Segoe UI", 10, "bold"), bg=CLR["accent"], fg="white",
                    relief="flat", cursor="hand2", padx=16, pady=6).pack(pady=20)

    def _reset_paramedis_pin_dialog(self):
        master_key = simpledialog.askstring("Reset PIN Paramedis", "Masukkan Master Key untuk mereset PIN Mode Paramedis:", show="*", parent=self.root)
        if master_key == "ITPSHE#2026":  # example master key, samakan dgn PIN admin
            CONFIG["paramedis_pin"] = _hash_pin("paramedis123")
            CONFIG["paramedis_pin_hashed"] = True
            try:
                _save_config(CONFIG)
                messagebox.showinfo("Berhasil", "PIN Mode Paramedis berhasil di-reset ke default: paramedis123", parent=self.root)
            except Exception as e:
                messagebox.showerror("Error", f"Gagal reset PIN: {e}", parent=self.root)
        elif master_key is not None:
            messagebox.showerror("Akses Ditolak", "Master Key salah!", parent=self.root)

    # ─────────────────────────────────────────────────────────────────
    # EXPORT
    # ─────────────────────────────────────────────────────────────────
    def _get_export_df(self):
        df = self._get_filtered_df().copy()
        if "avg_reaction" in df.columns:
            df["avg_reaction"] = pd.to_numeric(df["avg_reaction"], errors='coerce').round(1)

        def get_kesimpulan_sistol(sis):
            try:
                sis = float(sis)
                if sis < 100: return "Hipotensi"
                if sis < 140: return "Normal"
                if sis < 160: return "Hypertensi Ringan"
                if sis < 180: return "Hypertensi Sedang"
                return "Hypertensi Berat"
            except: return ""

        def get_kesimpulan_diastol(dia):
            try:
                dia = float(dia)
                if dia < 70: return "Hipotensi"
                if dia < 90: return "Normal"
                if dia < 100: return "Hypertensi Ringan"
                if dia < 120: return "Hypertensi Sedang"
                return "Hypertensi Berat"
            except: return ""

        def get_kesimpulan_nadi(nad):
            try:
                nad = float(nad)
                if nad < 60: return "Bradikardia"
                if nad < 100: return "Normal"
                if nad < 110: return "Takikardia Ringan"
                if nad < 120: return "Takikardia Sedang"
                return "Takikardia Berat"
            except: return ""

        def get_kesimpulan_alkohol(alk):
            try:
                alk = float(alk)
                if alk < 0.08: return "Lolos"
                return "Tidak Lolos"
            except: return ""

        if "kesp_sistol" in df.columns and "sistol" in df.columns:
            df["kesp_sistol"] = df.apply(lambda r: r["kesp_sistol"] if r.get("kesp_sistol") and str(r.get("kesp_sistol")).strip() not in ("", "-") else get_kesimpulan_sistol(r.get("sistol")), axis=1)
        if "kesp_diastol" in df.columns and "diastol" in df.columns:
            df["kesp_diastol"] = df.apply(lambda r: r["kesp_diastol"] if r.get("kesp_diastol") and str(r.get("kesp_diastol")).strip() not in ("", "-") else get_kesimpulan_diastol(r.get("diastol")), axis=1)
        if "kesp_nadi" in df.columns and "nadi" in df.columns:
            df["kesp_nadi"] = df.apply(lambda r: r["kesp_nadi"] if r.get("kesp_nadi") and str(r.get("kesp_nadi")).strip() not in ("", "-") else get_kesimpulan_nadi(r.get("nadi")), axis=1)
        if "kesimpulan_alkohol" in df.columns and "alcohol_test" in df.columns:
            df["kesimpulan_alkohol"] = df.apply(lambda r: r["kesimpulan_alkohol"] if r.get("kesimpulan_alkohol") and str(r.get("kesimpulan_alkohol")).strip() not in ("", "-") else get_kesimpulan_alkohol(r.get("alcohol_test")), axis=1)

        col_map = {
            "id": "No", "tanggal": "Tanggal", "nama": "Nama", "nik": "NIK", 
            "usia": "Usia", "jenis_kelamin": "Jenis Kelamin", "jabatan": "Bagian",
            "sistol": "TD Sistol", "diastol": "TD Diastol", "nadi": "Nadi", "alcohol_test": "Kadar Alkohol",
            "t1": "T1", "t2": "T2", "t3": "T3", "t4": "T4", "t5": "T5", "t6": "T6",
            "total": "Total HFT", "avg_reaction": "Rata-rata HFT",
            "fatigue_level": "Kesimpulan (Fatigue)",
            "kesp_sistol": "Kesimpulan (Sistol)",
            "kesp_diastol": "Kesimpulan (Diastol)",
            "kesp_nadi": "Kesimpulan (Nadi)",
            "keputusan": "Keputusan", "rekomendasi": "Rekomendasi",
            "keterangan": "Keterangan", "kesimpulan_alkohol": "Kesimpulan (Alkohol)"
        }
        df = df[[c for c in col_map.keys() if c in df.columns]]
        df.rename(columns=col_map, inplace=True)
        return df

    def export_excel(self):
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Export_FAST_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            title="Simpan Excel")
        if not dest:
            return
        try:
            df = self._get_export_df()
            df.to_excel(dest, index=False)
            wb     = load_workbook(dest)
            ws     = wb.active

            # Auto-fit kolom
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value or "")))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

            # Freeze baris pertama (header)
            ws.freeze_panes = "A2"

            # Style header
            header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Color maps for specific columns
            med_colors = {
                # Sistol & Diastol
                "Hipotensi":          ("D9E1F2", "1F4E78", False), # Soft Blue
                "Hypertensi Ringan":  ("2ECC71", "000000", False), # Hijau
                "Hypertensi Sedang":  ("F1C40F", "000000", False), # Kuning
                "Hypertensi Berat":   ("E74C3C", "FFFFFF", True),  # Merah

                # Nadi
                "Bradikardia":        ("D9E1F2", "1F4E78", False), # Soft Blue
                "Takikardia Ringan":  ("2ECC71", "000000", False), # Hijau
                "Takikardia Sedang":  ("F1C40F", "000000", False), # Kuning
                "Takikardia Berat":   ("E74C3C", "FFFFFF", True),  # Merah

                # Fatigue
                "Fatigue Ringan":     ("2ECC71", "000000", False), # Hijau
                "Fatigue Sedang":     ("F1C40F", "000000", False), # Kuning
                "Fatigue Berat":      ("E74C3C", "FFFFFF", True),  # Merah

                # Alkohol
                "Tidak Lolos":        ("E74C3C", "FFFFFF", True),  # Merah
            }

            col_idx = {}
            for col in range(1, ws.max_column + 1):
                val = ws.cell(1, col).value
                if val:
                    col_idx[val] = col

            fat_col = col_idx.get("Kesimpulan (Fatigue)")
            avg_col = col_idx.get("Rata-rata HFT")
            usia_col = col_idx.get("Usia")
            sistol_col = col_idx.get("TD Sistol")
            kes_sistol_col = col_idx.get("Kesimpulan (Sistol)")
            diastol_col = col_idx.get("TD Diastol")
            kes_diastol_col = col_idx.get("Kesimpulan (Diastol)")
            nadi_col = col_idx.get("Nadi")
            kes_nadi_col = col_idx.get("Kesimpulan (Nadi)")
            alkohol_col = col_idx.get("Kadar Alkohol")
            kes_alkohol_col = col_idx.get("Kesimpulan (Alkohol)")

            thr = CONFIG.get("age_warning_threshold", 60)
            for r in range(2, ws.max_row + 1):
                # 1. Usia warning (only highlight age cell)
                if usia_col:
                    try:
                        u_val = int(ws.cell(r, usia_col).value or 0)
                        if u_val >= thr:
                            ws.cell(r, usia_col).fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
                            ws.cell(r, usia_col).font = Font(color="922B21", bold=True)
                    except Exception:
                        pass

                # Helper to apply color to target cell
                def apply_cell_color(target_col, conclusion_val):
                    if target_col and conclusion_val in med_colors:
                        bg, fg, bld = med_colors[conclusion_val]
                        ws.cell(r, target_col).fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
                        ws.cell(r, target_col).font = Font(color=fg, bold=bld)

                # 2. Sistol
                if kes_sistol_col:
                    val = ws.cell(r, kes_sistol_col).value
                    apply_cell_color(sistol_col, val)
                    apply_cell_color(kes_sistol_col, val)

                # 3. Diastol
                if kes_diastol_col:
                    val = ws.cell(r, kes_diastol_col).value
                    apply_cell_color(diastol_col, val)
                    apply_cell_color(kes_diastol_col, val)

                # 4. Nadi
                if kes_nadi_col:
                    val = ws.cell(r, kes_nadi_col).value
                    apply_cell_color(nadi_col, val)
                    apply_cell_color(kes_nadi_col, val)

                # 5. Fatigue
                if fat_col:
                    val = ws.cell(r, fat_col).value
                    apply_cell_color(fat_col, val)
                    apply_cell_color(avg_col, val)

                # 6. Alkohol
                if kes_alkohol_col:
                    val = ws.cell(r, kes_alkohol_col).value
                    apply_cell_color(alkohol_col, val)
                    apply_cell_color(kes_alkohol_col, val)

                # Format reaction time columns (columns 7-15)
                for col in range(7, 15):
                    c = ws.cell(r, col)
                    if isinstance(c.value, (int, float)):
                        c.number_format = "0.0"
            wb.save(dest)
            self._update_status(f"✅ Excel disimpan: {os.path.basename(dest)}")
            if messagebox.askyesno("Berhasil",
                    f"File berhasil disimpan.\n{dest}\n\nBuka file sekarang?"):
                _open_file(dest)
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def export_csv(self):
        dest = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=f"Export_FAST_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            title="Simpan CSV")
        if not dest:
            return
        try:
            self._get_export_df().to_csv(dest, index=False, encoding="utf-8-sig")
            self._update_status(f"✅ CSV disimpan: {os.path.basename(dest)}")
            if messagebox.askyesno("Berhasil",
                    f"File berhasil disimpan.\n{dest}\n\nBuka file sekarang?"):
                _open_file(dest)
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    def export_pdf(self):
        if not PDF_AVAILABLE:
            messagebox.showerror(
                "Fitur Tidak Tersedia",
                "Export PDF tidak dapat digunakan pada versi ini.\n"
                "Gunakan Export Excel atau CSV sebagai alternatif.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            initialfile=f"Export_FAST_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            title="Simpan PDF")
        if not dest:
            return
        try:
            df      = self._get_filtered_df()
            company = CONFIG.get("company_name", "PT. NAMA PERUSAHAAN")
            pdf     = FPDF(orientation="L")
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(277, 8, company, ln=1, align="C")
            pdf.set_font("Arial", "B", 11)
            pdf.cell(277, 7, "Laporan Hasil Tes FAST", ln=1, align="C")
            pdf.set_font("Arial", "", 8)
            pdf.cell(
                277, 6,
                f"Tanggal Cetak: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Dibuat oleh: {CONFIG.get('created_by', 'Marsellinus, Khilqi, Idris')}",
                ln=1, align="C")
            pdf.ln(3)

            hdrs = ["Tanggal", "Nama", "NIK", "Usia", "Divisi",
                    "T1", "T2", "T3", "T4", "T5", "T6",
                    "Total", "Avg", "Fatigue"]
            cws  = [38, 38, 30, 14, 28,
                    13, 13, 13, 13, 13, 13, 16, 16, 32]
            pdf.set_font("Arial", "B", 8)
            pdf.set_fill_color(52, 152, 219)
            pdf.set_text_color(255, 255, 255)
            for h, w in zip(hdrs, cws):
                pdf.cell(w, 8, h, border=1, align="C", fill=True)
            pdf.ln()

            pdf.set_font("Arial", "", 7)
            fill_colors = {
                "Fatigue Ringan": (213, 245, 227),
                "Fatigue Sedang": (254, 249, 231),
                "Fatigue Berat":  (250, 219, 216),
            }
            thr = CONFIG.get("age_warning_threshold", 60)
            for _, row in df.iterrows():
                fat = str(row.get("fatigue_level", ""))
                try:
                    age_v = int(row.get("usia", 0) or 0)
                except Exception:
                    age_v = 0
                if age_v > thr:
                    pdf.set_fill_color(250, 219, 216)
                    pdf.set_text_color(146, 43, 33)
                    fill = True
                elif fat in fill_colors:
                    r2, g2, b2 = fill_colors[fat]
                    pdf.set_fill_color(r2, g2, b2)
                    pdf.set_text_color(44, 62, 80)
                    fill = True
                else:
                    pdf.set_fill_color(255, 255, 255)
                    pdf.set_text_color(44, 62, 80)
                    fill = False
                row_vals = [
                    str(row.get("tanggal",      ""))[:16],
                    str(row.get("nama",         ""))[:16],
                    str(row.get("nik",          "")),
                    str(row.get("usia",         "")),
                    str(row.get("divisi",       ""))[:12],
                    f"{row.get('t1', 0):.0f}",
                    f"{row.get('t2', 0):.0f}",
                    f"{row.get('t3', 0):.0f}",
                    f"{row.get('t4', 0):.0f}",
                    f"{row.get('t5', 0):.0f}",
                    f"{row.get('t6', 0):.0f}",
                    f"{row.get('total',        0):.0f}",
                    f"{row.get('avg_reaction', 0):.1f}",
                    fat[:14],
                ]
                for val, w in zip(row_vals, cws):
                    pdf.cell(w, 7, val, border=1, align="C", fill=fill)
                pdf.ln()
            pdf.output(dest)
            self._update_status(f"PDF: {os.path.basename(dest)}")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    # ─────────────────────────────────────────────────────────────────
    # IMPORT RIWAYAT TES
    # ─────────────────────────────────────────────────────────────────
    def import_data(self):
        path = filedialog.askopenfilename(
            title="Pilih File Riwayat Tes",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"),
                       ("Semua",     "*.*")])
        if not path:
            return
        try:
            df = (pd.read_csv(path)
                  if path.lower().endswith(".csv")
                  else pd.read_excel(path))
                  
            df.columns = [str(c).lower().strip() for c in df.columns]
            rev_map = {
                "tanggal": "tanggal", "nama": "nama", "nik": "nik",
                "usia": "usia", "jenis kelamin": "jenis_kelamin", "jabatan": "jabatan",
                "td sistol": "td_sistol", "td diastol": "td_diastol", "nadi": "nadi", "kadar alkohol": "alcohol_test",
                "t1": "t1", "t2": "t2", "t3": "t3", "t4": "t4", "t5": "t5", "t6": "t6",
                "rata-rata hft": "avg_reaction",
                "kesimpulan (fatigue)": "fatigue_level",
                "kesimpulan (sistol)": "kesimpulan_sistol",
                "kesimpulan (diastol)": "kesimpulan_diastol",
                "kesimpulan (nadi)": "kesimpulan_nadi",
                "kesimpulan (alkohol)": "kesimpulan_alkohol",
                "keputusan": "keputusan", "rekomendasi": "rekomendasi", "keterangan": "keterangan",
                "jenis_kelamin": "jenis_kelamin", "td_sistol": "td_sistol", "td_diastol": "td_diastol",
                "alcohol_test": "alcohol_test", "avg_reaction": "avg_reaction", "fatigue_level": "fatigue_level",
                "kesimpulan_sistol": "kesimpulan_sistol", "kesimpulan_diastol": "kesimpulan_diastol",
                "kesimpulan_nadi": "kesimpulan_nadi", "kesimpulan_alkohol": "kesimpulan_alkohol",
                "perusahaan/bagian": "divisi", "divisi": "divisi"
            }
            df.rename(columns=rev_map, inplace=True)
            df = df.where(pd.notnull(df), None)
            
            with sqlite3.connect(DB_PATH) as conn:
                cur     = conn.execute("SELECT * FROM results LIMIT 0")
                db_cols = [d[0] for d in cur.description if d[0] != "id"]

                required = {"tanggal", "nama", "nik"}
                missing  = required - set(df.columns)
                if missing:
                    messagebox.showerror("Format Salah",
                        f"File tidak memiliki kolom wajib: {', '.join(missing)}")
                    return

                existing = set()
                for row_ex in conn.execute("SELECT nik, tanggal, avg_reaction, t1, t2 FROM results"):
                    nik_ex = str(row_ex[0] or "").strip().upper()
                    tgl_ex = str(row_ex[1] or "").strip()
                    try: avg_ex = round(float(row_ex[2] or 0), 2)
                    except Exception: avg_ex = 0.0
                    try: t1_ex = round(float(row_ex[3] or 0), 2)
                    except Exception: t1_ex = 0.0
                    try: t2_ex = round(float(row_ex[4] or 0), 2)
                    except Exception: t2_ex = 0.0
                    existing.add((nik_ex, tgl_ex, avg_ex, t1_ex, t2_ex))

                n = skipped = 0
                for _, row in df.iterrows():
                    nik_val = str(row.get("nik") or "").strip().upper()
                    tgl_val = str(row.get("tanggal") or "").strip()
                    try: avg_val = round(float(row.get("avg_reaction") or row.get("rata-rata hft") or 0), 2)
                    except Exception: avg_val = 0.0
                    try: t1_val = round(float(row.get("t1") or 0), 2)
                    except Exception: t1_val = 0.0
                    try: t2_val = round(float(row.get("t2") or 0), 2)
                    except Exception: t2_val = 0.0

                    # Key unik per tes (memungkinkan 3x tes per hari untuk NIK yang sama)
                    key = (nik_val, tgl_val, avg_val, t1_val, t2_val)
                    if key in existing:
                        skipped += 1
                        continue

                    rd = {}
                    for col in db_cols:
                        val = row.get(col)
                        if col in ("t1", "t2", "t3", "t4", "t5", "t6",
                                   "avg_reaction") and val is not None:
                            try:
                                val = float(val)
                                if col != "avg_reaction" and not (0 <= val <= 5000):
                                    val = 0.0
                            except Exception:
                                val = 0.0
                        rd[col] = val if val is not None else "-"
                    ph = ",".join(["?"] * len(db_cols))
                    conn.execute(
                        f"INSERT INTO results ({','.join(db_cols)}) VALUES ({ph})",
                        tuple(rd[c] for c in db_cols))
                    existing.add(key)
                    n += 1
                conn.commit()

            self.refresh_dashboard()
            msg = f"{n} baris berhasil diimpor."
            if skipped:
                msg += f" {skipped} baris dilewati (duplikat)."
            messagebox.showinfo("Berhasil", msg)
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    # ─────────────────────────────────────────────────────────────────
    # ─────────────────────────────────────────────────────────────────
    # CONTEXT MENU (KLIK KANAN)
    # ─────────────────────────────────────────────────────────────────
    def _bind_context_menus(self):
        if hasattr(self, 'tree'): self.tree.bind("<Button-3>", lambda e: self._show_context_menu(e, self.tree, "main"))
        if hasattr(self, 'tree_filter'): self.tree_filter.bind("<Button-3>", lambda e: self._show_context_menu(e, self.tree_filter, "filter"))
        if hasattr(self, 'tree_master'): self.tree_master.bind("<Button-3>", lambda e: self._show_context_menu(e, self.tree_master, "karyawan"))
        if hasattr(self, 'tree_master_kontraktor'): self.tree_master_kontraktor.bind("<Button-3>", lambda e: self._show_context_menu(e, self.tree_master_kontraktor, "kontraktor"))
        if hasattr(self, 'tree_medis'): self.tree_medis.bind("<Button-3>", lambda e: self._show_context_menu(e, self.tree_medis, "medis"))

    def _show_context_menu(self, event, tree, tree_type):
        iid = tree.identify_row(event.y)
        if not iid:
            return
        tree.selection_set(iid)
        
        menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 9), bg=CLR["bg_panel"], fg=CLR["text_main"])
        
        if tree_type in ["main", "filter", "medis"]:
            menu.add_command(label="📄 Lihat Riwayat Pasien", command=lambda: self._open_history_from_tree(tree))
            medis_label = "🩺 Input Alcohol Test" if self.app_mode_var.get() == "SHE Officer" else "🩺 Input Data Medis"
            menu.add_command(label=medis_label, command=lambda: self._open_medis_dialog(from_medis_tab=(tree_type=="medis")))
            menu.add_separator()
            menu.add_command(label="📋 Copy NIK", command=lambda: self._copy_nik(tree, 3))
            menu.add_separator()
            menu.add_command(label="🗑️ Hapus Data Tes", command=lambda: self._delete_from_context(tree))
            
        elif tree_type in ["karyawan", "kontraktor"]:
            is_kontraktor = (tree_type == "kontraktor")
            menu.add_command(label="🎯 Pilih untuk Tes", command=lambda: self._select_master_for_test(tree))
            menu.add_separator()
            menu.add_command(label="✏️ Edit Data", command=lambda: self._mk_edit(is_kontraktor))
            menu.add_command(label="📋 Copy NIK", command=lambda: self._copy_nik(tree, 2))
            menu.add_separator()
            menu.add_command(label="🗑️ Hapus Data", command=lambda: self._mk_delete(is_kontraktor))

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _copy_nik(self, tree, col_idx):
        sel = tree.selection()
        if sel:
            v = tree.item(sel[0], "values")
            if len(v) > col_idx:
                nik = str(v[col_idx]).strip()
                self.root.clipboard_clear()
                self.root.clipboard_append(nik)
                self.root.update()
                self._update_status(f"NIK disalin: {nik}")

    def _delete_from_context(self, tree):
        sel = tree.selection()
        if not sel: return
        
        # Ambil ID database dari GUI SEBELUM dialog muncul
        to_delete_ids = []
        for item in sel:
            v = tree.item(item, "values")
            to_delete_ids.append(v[0])

        pin = simpledialog.askstring("Admin PIN", "Masukkan PIN Admin:", show="*", parent=self.root)
        if not _verify_pin(pin or ""):
            messagebox.showerror("Akses Ditolak", "PIN Admin salah!", parent=self.root)
            return
        if not messagebox.askyesno("Konfirmasi", f"Hapus data terpilih?"):
            return
        try:
            with sqlite3.connect(DB_PATH) as conn:
                c = conn.cursor()
                for db_id in to_delete_ids:
                    c.execute("DELETE FROM results WHERE id=?", (db_id,))
                conn.commit()
            self.refresh_dashboard()
            try: self.update_table()
            except: pass
            try: self.update_table_medis()
            except: pass
            try: self._apply_filter()
            except: pass
            messagebox.showinfo("Berhasil", "Data berhasil dihapus.")
        except Exception as e:
            messagebox.showerror("Gagal", str(e))

    # ─────────────────────────────────────────────────────────────────
    # APP CLOSE GUARD
    def _on_app_close(self):
        """Konfirmasi sebelum tutup app jika tes sedang berjalan."""
        if self.active_test_window is not None:
            if not messagebox.askyesno(
                "Tes Sedang Berlangsung",
                "Ada pengujian FAST yang sedang aktif.\n"
                "Data trial yang belum selesai akan hilang.\n\n"
                "Tutup aplikasi sekarang?"):
                return
            try:
                self.active_test_window.destroy()
            except Exception:
                pass
            self.active_test_window = None
        self.listener_running = False
        self.root.destroy()

    # ─────────────────────────────────────────────────────────────────
    # UTILS
    # ─────────────────────────────────────────────────────────────────
    def _update_status(self, msg):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        self.status_bar_var.set(f"[{ts}] {msg}")
        self.root.after(6000, lambda: self.status_bar_var.set("Siap."))


# ==========================================
# ENTRY POINT
# ==========================================
if __name__ == "__main__":
    try:
        init_db()
        init_excel()
        _migrate_pin_to_hash()   # upgrade plain PIN → hashed
        _migrate_paramedis_pin_to_hash()   # upgrade plain PIN Paramedis → hashed
        root = tk.Tk()
        
        # Set Application Icon
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "assets", "ratig.ico")
        else:
            icon_path = os.path.join(ASSETS_DIR, "ratig.ico")
            
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(icon_path)
            except Exception as e:
                logging.warning(f"Gagal memuat icon: {e}")
                
        app  = RatigApp(root)
        root.mainloop()
    except Exception as _startup_err:
        import traceback
        _msg = (
            f"FAST gagal dijalankan.\n\n"
            f"{type(_startup_err).__name__}: {_startup_err}\n\n"
            f"Detail error tersimpan di:\n{LOG_FILE}"
        )
        logging.critical(f"Startup crash: {traceback.format_exc()}") 
        try:
            # coba tampilkan dialog jika tkinter sudah terinisialisasi
            import tkinter.messagebox as _mb
            _mb.showerror("FAST - Error Startup", _msg)
        except Exception:
            pass