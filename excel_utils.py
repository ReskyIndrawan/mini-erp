import os
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

TITLE = "不具合品一覧表"
EXCEL_NAME = "不具合品一覧表.xlsx"


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


def append_excel(folder, rowdata, creator):
    filepath = create_excel_if_not_exists(folder, creator)
    wb = load_workbook(filepath)
    ws = wb.active
    ws.append(rowdata)
    wb.save(filepath)
    return filepath
