import os
import datetime
import json
import sys
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import urllib.parse
import subprocess
import codecs
import re

TITLE = "不具合品一覧表"
EXCEL_NAME = "不具合品一覧表.xlsx"

# History management constants
APP_NAME = "defect_data_app"
HISTORY_FILENAME = "recent_excel_files.json"
MAX_HISTORY = 10


def get_config_dir():
    """Get cross-platform config directory"""
    if sys.platform.startswith("win"):
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / APP_NAME
    # Prefer XDG_CONFIG_HOME if set, otherwise ~/.config
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        return Path(xdg) / APP_NAME
    return Path.home() / ".config" / APP_NAME


class ExcelHistoryManager:
    """
    Menyimpan history path file dalam format tampilan Jepang (separator '¥').
    Saat mengambil dan menampilkan ke UI, gunakan apa adanya.
    Saat membuka file, gunakan open_file_safely() yang akan konversi ke Windows style.
    """

    def __init__(self, max_items=MAX_HISTORY):
        self.max_items = max_items
        self.config_dir = get_config_dir()
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.path = self.config_dir / HISTORY_FILENAME
        self._items = self._load()

    def _load(self):
        """Load history from JSON file (sudah disimpan dalam format tampilan '¥')."""
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return [str(x) for x in data if isinstance(x, str)]
            except Exception:
                pass
        return []

    def save(self):
        """Save history to JSON file"""
        try:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self._items, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print("Failed saving recent files:", e)

    def add(self, filepath):
        """
        Add file to history.
        Simpan path apa adanya tanpa konversi.
        """
        raw = str(filepath).strip()

        if raw in self._items:
            self._items.remove(raw)
        self._items.insert(0, raw)
        self._items = self._items[: self.max_items]
        self.save()

    def clear(self):
        """Clear all history"""
        self._items = []
        try:
            if self.path.exists():
                self.path.unlink()
        except Exception:
            pass

    def items(self):
        """Get all history items (dalam format tampilan '¥')."""
        return list(self._items)

    def remove(self, filepath):
        """Remove specific file from history."""
        raw = str(filepath).strip()
        if raw in self._items:
            self._items.remove(raw)
            self.save()


def create_excel_if_not_exists(folder, creator):
    filepath = os.path.join(folder, EXCEL_NAME)
    if not os.path.exists(filepath):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # Title
        ws.merge_cells("A1:L1")
        ws["A1"] = TITLE
        ws["A1"].font = Font(size=16, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        # Date & creator
        today = datetime.date.today()
        ws["A2"] = f"月間期間: {today.year}-{today.month}"
        ws["C2"] = f"作成者: {creator}"

        # Header
        headers = [
            "発生月",
            "累計",
            "№",
            "発生日",
            "項目",
            "事象",
            "事象（一次）",
            "事象（二次）",
            "品番",
            "サプライヤー名",
            "不良発生連絡書発行",
            "不良発生№",
        ]
        ws.append(headers)

        # Styling header
        header_fill = PatternFill("solid", fgColor="C0C0C0")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=3, column=col)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill
            cell.border = border
            cell.font = Font(bold=True)

        wb.save(filepath)
    return filepath


def normalize_japanese_path(raw_path: str) -> str:
    """
    Normalize path yang mungkin datang dari search-ms: URI atau URL-encoded.
    - Decode URL (%xx)
    - Ganti '/' menjadi '\\' untuk Windows style (sementara, sebelum tampilan)
    - Perbaiki UNC '//' -> '\\\\'
    """
    if not raw_path:
        return raw_path

    s = raw_path.strip()

    # Jika path diawali dengan 'search-ms:', coba ambil crumb=location:
    if s.lower().startswith("search-ms:"):
        crumb_prefix = "crumb=location:"
        idx = s.lower().find(crumb_prefix)
        if idx != -1:
            encoded_path = s[idx + len(crumb_prefix) :]
            s = urllib.parse.unquote(encoded_path)
        # jika tidak ditemukan, biarkan s apa adanya

    # Decode bentuk $uXXXX menjadi unicode nyata (contoh: $u3000)
    s = _decode_dollar_u_sequences(s)
    # Decode bentuk \uXXXX jika kebetulan ada
    try:
        s = codecs.decode(s, "unicode_escape")
    except Exception:
        pass

    # Untuk konsistensi internal Windows: '/' -> '\'
    s = s.replace("/", "\\")

    # Pastikan UNC benar: '//' atau '\\' di-normalisasi ke '\\\\' diawal
    if s.startswith("//"):
        s = "\\" + s  # '//' -> '\//', lalu replace di bawah
    if s.startswith("\\/"):
        s = s.replace("\\/", "\\\\", 1)
    if (
        s.startswith("\\\\") is False
        and s.startswith("\\")
        and len(s) > 1
        and s[1] == "\\"
    ):
        # Sudah \\X... tidak perlu
        pass

    # Bersihkan ganda campuran '/': ganti '\/' -> '\'
    s = s.replace("\\/", "\\")
    # Tidak gunakan Path.resolve agar tidak memaksa backslash saat menyimpan history
    return s


def _decode_dollar_u_sequences(s: str) -> str:
    """
    Ubah literal seperti $u3000 menjadi karakter unicode sebenarnya (U+3000).
    Juga dukung $U+XXXX.
    """
    if not s:
        return s

    def repl(m):
        hexpart = m.group(1) or m.group(2)
        try:
            codepoint = int(hexpart, 16)
            return chr(codepoint)
        except Exception:
            return m.group(0)

    # Pola $uXXXX atau $UXXXX atau $U+XXXX
    return re.sub(r"\$u([0-9a-fA-F]{4})|\$U\+?([0-9a-fA-F]{4})", repl, s)


def convert_path_to_display_style(path: str) -> str:
    """
    Ubah path ke format tampilan Jepang:
    - Semua separator ('\\' dan '/') -> '¥'
    - UNC leading: '\\\\server\\share' -> '¥¥server¥share'
    """
    if not path:
        return path
    # Normalisasi dulu agar konsisten
    s = normalize_japanese_path(path)
    # Ganti semua backslash menjadi '¥'
    s = s.replace("\\\\", "¥¥")  # jaga UNC double
    s = s.replace("\\", "¥")
    # Ganti sisa '/' (jika ada) menjadi '¥'
    s = s.replace("/", "¥")
    return s


def convert_path_to_windows_style(path: str) -> str:
    """
    Konversi dari tampilan Jepang ke Windows path:
    - '¥' -> '\\'
    - '/' -> '\\'
    - Perbaiki UNC menjadi '\\\\server\\share'
    """
    if not path:
        return path

    s = path.strip()

    # Decode $uXXXX / \uXXXX jika ada
    s = _decode_dollar_u_sequences(s)
    try:
        s = codecs.decode(s, "unicode_escape")
    except Exception:
        pass

    # Tampilan mungkin sudah punya '¥'
    s = s.replace("¥", "\\")
    s = s.replace("/", "\\")

    # Perbaiki variasi UNC yang mungkin muncul
    if s.startswith("//"):
        s = "\\" + s  # jadi '\//...'
    s = s.replace("\\/", "\\")
    if s.startswith("\\") and not s.startswith("\\\\"):
        # Jika dimulai satu backslash saja dan berikutnya backslash, biarkan.
        # Jika dimulai satu backslash saja sebelum server name, ubah ke dua.
        parts = s.split("\\")
        if len(parts) > 2 and parts[1] and parts[2]:
            s = "\\" + s  # jadikan UNC
    return s


def open_file_safely(file_path: str):
    """
    Buka file dengan aman.
    Gunakan path apa adanya tanpa konversi.
    """
    if not file_path:
        raise ValueError("Path file tidak boleh kosong.")

    # Periksa keberadaan file langsung
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File tidak ditemukan: {file_path}")

    try:
        if os.name == "nt":  # Windows
            os.startfile(file_path)
        elif os.name == "posix":  # macOS/Linux
            opener = "open" if "darwin" in os.uname().sysname.lower() else "xdg-open"
            subprocess.call([opener, file_path])
    except Exception as e:
        raise RuntimeError(f"Gagal membuka file: {e}")


def append_excel(folder, rowdata, creator):
    filepath = create_excel_if_not_exists(folder, creator)
    wb = load_workbook(filepath)
    ws = wb.active
    ws.append(rowdata)
    wb.save(filepath)
    return filepath
