import os
import datetime
import json
import sys
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import urllib.parse

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
    def __init__(self, max_items=MAX_HISTORY):
        self.max_items = max_items
        self.config_dir = get_config_dir()
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.path = self.config_dir / HISTORY_FILENAME
        self._items = self._load()

    def _load(self):
        """Load history from JSON file"""
        if self.path.exists():
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # ensure it's a list of strings
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
        """Add file to history"""
        filepath = str(Path(filepath).resolve())
        if filepath in self._items:
            self._items.remove(filepath)
        self._items.insert(0, filepath)
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
        """Get all history items"""
        return list(self._items)

    def remove(self, filepath):
        """Remove specific file from history"""
        f = str(Path(filepath).resolve())
        if f in self._items:
            self._items.remove(f)
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
    Normalize Windows path that may contain Japanese characters encoded in search-ms: URI format.
    Converts search-ms: URI to normal Windows path string.
    If input is already a normal path, returns as is.
    """
    if not raw_path:
        return raw_path

    raw_path = raw_path.strip()

    # Jika path diawali dengan 'search-ms:', parse dan decode
    if raw_path.lower().startswith("search-ms:"):
        # Contoh format:
        # search-ms:displayname=...&crumb=location:<encoded_path>
        # Kita cari crumb=location: dan ambil sisanya

        # Cari 'crumb=location:' di string
        crumb_prefix = "crumb=location:"
        idx = raw_path.lower().find(crumb_prefix)
        if idx != -1:
            # Ambil substring setelah crumb=location:
            encoded_path = raw_path[idx + len(crumb_prefix) :]
            # encoded_path bisa mengandung karakter %xx, decode URL
            decoded_path = urllib.parse.unquote(encoded_path)

            # Windows path biasanya diawali dengan \\server\share
            # Pastikan diawali dengan backslash
            if decoded_path.startswith("\\\\") or decoded_path.startswith("/"):
                # Ganti / dengan \ jika ada
                normalized_path = decoded_path.replace("/", "\\")
                return normalized_path
            else:
                # Jika tidak sesuai, kembalikan hasil decode apa adanya
                return decoded_path

        else:
            # Jika crumb=location: tidak ditemukan, kembalikan apa adanya
            return raw_path

    else:
        # Jika bukan search-ms, kembalikan apa adanya
        return raw_path


def escape_path_for_japanese_locale(path: str) -> str:
    """
    Escape path ke format Unicode escaped agar kompatibel dengan sistem Jepang.
    Contoh: "C:\\テスト" -> "\\u0043\\u003a\\u005c\\u30c6\\u30b9\\u30c8"
    """
    if not path:
        return path
    # Escape semua karakter ke Unicode
    return "".join(f"\\u{ord(c):04x}" for c in path)


def unescape_path_for_japanese_locale(escaped_path: str) -> str:
    """
    Unescape path dari format Unicode escaped ke bentuk asli.
    Contoh: "\\u0043\\u003a\\u005c\\u30c6\\u30b9\\u30c8" -> "C:\\テスト"
    """
    if not escaped_path:
        return escaped_path
    # Unescape \\uXXXX menjadi karakter asli
    import codecs

    try:
        return codecs.decode(escaped_path, "unicode_escape")
    except Exception:
        return escaped_path  # Jika gagal, kembalikan apa adanya


def append_excel(folder, rowdata, creator):
    filepath = create_excel_if_not_exists(folder, creator)
    wb = load_workbook(filepath)
    ws = wb.active
    ws.append(rowdata)
    wb.save(filepath)
    return filepath
