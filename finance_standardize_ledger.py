# 示例说明：MODULES 是「宽表分栏 -> 规范流水账」的示例映射，
# 列号与名称均为虚构示例，使用时请按自己的表格结构替换。

from __future__ import annotations

import json
import re
import sys
from copy import copy
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.views import Selection


INPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ledger.xlsx")
OUTPUT_DIR = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("output")
OUTPUT = OUTPUT_DIR / "财务规范版.xlsx"
REPORT = OUTPUT_DIR / "finance_standard_report.json"

SOURCE_SHEET = "Sheet1"
LEDGER_SHEET = "规范流水账"

MODULES = [
    {"name": "甲方付款", "category": "回款/开票", "cols": (1, 3), "date": 1, "main": 2, "note": 3, "fill": "EAF3F8"},
    {"name": "成本票及税金", "category": "成本/税金", "cols": (4, 8), "date": 4, "main": 5, "tax": 6, "tax_cost": 7, "party": 8, "note": None, "fill": "FFF2CC"},
    {"name": "内部转入", "category": "内部往来", "cols": (9, 11), "date": 9, "main": 10, "note": 11, "fill": "E2F0D9"},
    {"name": "内部转出", "category": "内部往来", "cols": (12, 14), "date": 12, "main": 13, "note": 14, "fill": "E2F0D9"},
    {"name": "个人往来-甲", "category": "个人往来", "cols": (15, 17), "date": 15, "main": 16, "note": 17, "fill": "EDE7F6"},
    {"name": "个人往来-乙", "category": "个人往来", "cols": (18, 20), "date": 18, "main": 19, "note": 20, "fill": "EDE7F6"},
    {"name": "零星费用-丙", "category": "零星费用", "cols": (21, 23), "date": 21, "main": 22, "note": 23, "fill": "FCE4D6"},
    {"name": "粘接砂浆-供应商A", "category": "材料采购", "cols": (24, 26), "date": 24, "main": 25, "note": 26, "fill": "E2F0D9"},
    {"name": "网格布钉-供应商B", "category": "材料采购", "cols": (27, 29), "date": 27, "main": 28, "note": 29, "fill": "E2F0D9"},
    {"name": "网格布钉-供应商C", "category": "材料采购", "cols": (30, 32), "date": 30, "main": 31, "note": 32, "fill": "E2F0D9"},
    {"name": "真石漆-品牌D", "category": "材料采购", "cols": (33, 35), "date": 33, "main": 34, "note": 35, "fill": "E2F0D9"},
    {"name": "现场零星材料", "category": "零星材料", "cols": (36, 38), "date": 36, "main": 37, "note": 38, "fill": "FCE4D6"},
    {"name": "施工队-甲", "category": "人工劳务", "cols": (39, 41), "date": 39, "main": 40, "note": 41, "fill": "DDEBF7"},
    {"name": "施工队-乙", "category": "人工劳务", "cols": (42, 44), "date": 42, "main": 43, "note": 44, "fill": "DDEBF7"},
    {"name": "吊篮曲臂车", "category": "机械租赁", "cols": (45, 47), "date": 45, "main": 46, "note": 47, "fill": "D9EAD3"},
    {"name": "其他开支", "category": "其他支出", "cols": (48, 50), "date": 48, "main": 49, "note": 50, "fill": "E7E6E6"},
    {"name": "未报/补充", "category": "待核事项", "cols": (51, 52), "date": None, "main": 51, "note": 52, "fill": "F4CCCC"},
]


def snapshot_source(ws):
    return {
        "max_row": ws.max_row,
        "max_column": ws.max_column,
        "merged_ranges": [str(rng) for rng in ws.merged_cells.ranges],
        "values": [
            [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(1, ws.max_row + 1)
        ],
    }


def normalize_date(value):
    if value in (None, ""):
        return None
    if isinstance(value, date):
        return value
    text = str(value).strip()
    m = re.fullmatch(r"(\d{2})\.(\d{1,2})\.(\d{1,2})", text)
    if not m:
        return text
    yy, mm, dd = map(int, m.groups())
    year = 2000 + yy
    try:
        return date(year, mm, dd)
    except ValueError:
        return text


def cell_ref(sheet_name, row, col):
    return f"='{sheet_name}'!{get_column_letter(col)}{row}"


def group_has_data(ws, row, start_col, end_col):
    return any(ws.cell(row, c).value is not None for c in range(start_col, end_col + 1))


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(INPUT, data_only=False)
    source = wb[SOURCE_SHEET]
    before_source = snapshot_source(source)

    if LEDGER_SHEET in wb.sheetnames:
        del wb[LEDGER_SHEET]
    ledger = wb.create_sheet(LEDGER_SHEET, 0)

    dark = "17365D"
    header = "1F4E78"
    subheader = "5B9BD5"
    total = "D9EAF7"
    summary = "EAF3F8"
    white = "FFFFFF"
    text = "1F2933"
    gray = "D9E2EC"
    deep_gray = "7F8C8D"

    thin = Side(style="thin", color=gray)
    medium = Side(style="medium", color="8FA7BD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    top_border = Border(left=thin, right=thin, top=medium, bottom=thin)

    base_font = Font(name="Microsoft YaHei", size=10, color=text)
    small_font = Font(name="Microsoft YaHei", size=9, color=text)
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color=white)
    title_font = Font(name="Microsoft YaHei", size=16, bold=True, color=white)
    label_font = Font(name="Microsoft YaHei", size=10, bold=True, color=text)

    # Source sheet: financial style without changing original values/formulas.
    source.sheet_view.showGridLines = False
    source.freeze_panes = None
    source.sheet_view.topLeftCell = "A1"
    source.sheet_view.pane = None
    source.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    for row in source.iter_rows(min_row=1, max_row=84, min_col=1, max_col=52):
        for cell in row:
            cell.font = copy(base_font)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            if cell.value is None:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.number_format = '#,##0.00;[Red]-#,##0.00;"-"'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    for module in MODULES:
        fill = PatternFill("solid", fgColor=module["fill"])
        for c in range(module["cols"][0], module["cols"][1] + 1):
            for r in range(3, 80):
                cell = source.cell(r, c)
                if cell.value is not None:
                    cell.fill = fill
            for r in (2, 81):
                cell = source.cell(r, c)
                cell.fill = PatternFill("solid", fgColor=header if r == 2 else subheader)
                cell.font = copy(header_font)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(left=thin, right=thin, top=medium, bottom=medium)
            source.cell(1, c).border = top_border

    for r in range(80, 85):
        for c in range(1, 53):
            cell = source.cell(r, c)
            cell.fill = PatternFill("solid", fgColor=total if r in (80, 82, 83, 84) else subheader)
            cell.font = copy(header_font if r == 81 else label_font)
            cell.border = top_border if r == 80 else border

    for cell in source[1]:
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = copy(title_font if cell.coordinate == "B1" else header_font)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    source.row_dimensions[1].height = 34
    source.row_dimensions[2].height = 30
    source.row_dimensions[81].height = 30
    for r in range(3, 80):
        source.row_dimensions[r].height = 22
    for r in range(80, 85):
        source.row_dimensions[r].height = 25

    date_cols = {m["date"] for m in MODULES if m.get("date")}
    note_cols = {m["note"] for m in MODULES if m.get("note")}
    for c in range(1, 53):
        letter = get_column_letter(c)
        if c in date_cols:
            source.column_dimensions[letter].width = 11
        elif c in note_cols:
            source.column_dimensions[letter].width = 20
            for r in range(3, 80):
                source.cell(r, c).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            source.column_dimensions[letter].width = 14

    # Normalized ledger sheet.
    title = "示例项目 - 财务规范流水账"
    ledger.merge_cells("A1:L1")
    ledger["A1"] = title
    ledger["A1"].fill = PatternFill("solid", fgColor=dark)
    ledger["A1"].font = copy(title_font)
    ledger["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ledger.row_dimensions[1].height = 34

    headers = ["序号", "业务大类", "明细分类", "原表行", "日期", "往来对象/单位", "项目", "金额/内容", "税金", "税金成本", "备注", "源位置"]
    for idx, value in enumerate(headers, 1):
        cell = ledger.cell(2, idx, value)
        cell.fill = PatternFill("solid", fgColor=header)
        cell.font = copy(header_font)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=medium, bottom=medium)
    ledger.row_dimensions[2].height = 28

    rows = []
    serial = 1
    for module in MODULES:
        start_col, end_col = module["cols"]
        for row_idx in range(3, 80):
            if not group_has_data(source, row_idx, start_col, end_col):
                continue
            main_col = module["main"]
            note_col = module.get("note")
            tax_col = module.get("tax")
            tax_cost_col = module.get("tax_cost")
            party_col = module.get("party")
            date_col = module.get("date")
            src_cells = f"{get_column_letter(start_col)}{row_idx}:{get_column_letter(end_col)}{row_idx}"
            rows.append(
                {
                    "serial": serial,
                    "category": module["category"],
                    "module": module["name"],
                    "source_row": row_idx,
                    "date_value": normalize_date(source.cell(row_idx, date_col).value) if date_col else None,
                    "party_formula": cell_ref(SOURCE_SHEET, row_idx, party_col) if party_col else "",
                    "item": source.cell(2, main_col).value or module["name"],
                    "main_formula": cell_ref(SOURCE_SHEET, row_idx, main_col),
                    "tax_formula": cell_ref(SOURCE_SHEET, row_idx, tax_col) if tax_col else "",
                    "tax_cost_formula": cell_ref(SOURCE_SHEET, row_idx, tax_cost_col) if tax_cost_col else "",
                    "note_formula": cell_ref(SOURCE_SHEET, row_idx, note_col) if note_col else "",
                    "source_ref": src_cells,
                    "fill": module["fill"],
                }
            )
            serial += 1

    for out_row, item in enumerate(rows, 3):
        values = [
            item["serial"],
            item["category"],
            item["module"],
            item["source_row"],
            item["date_value"],
            item["party_formula"],
            item["item"],
            item["main_formula"],
            item["tax_formula"],
            item["tax_cost_formula"],
            item["note_formula"],
            item["source_ref"],
        ]
        for c, value in enumerate(values, 1):
            cell = ledger.cell(out_row, c)
            if isinstance(value, str) and value.startswith("="):
                cell.value = value
            else:
                cell.value = value
            cell.font = copy(small_font)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c in (2, 3):
                cell.fill = PatternFill("solid", fgColor=item["fill"])
            elif out_row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
            if c in (6, 7, 11, 12):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c == 5:
                cell.number_format = "yyyy-mm-dd"
            if c in (8, 9, 10):
                cell.number_format = '#,##0.00;[Red]-#,##0.00;"-"'
                cell.alignment = Alignment(horizontal="right", vertical="center")

    end_row = max(3, len(rows) + 2)
    widths = [8, 13, 16, 9, 13, 16, 16, 14, 13, 13, 28, 14]
    for idx, width in enumerate(widths, 1):
        ledger.column_dimensions[get_column_letter(idx)].width = width
    for r in range(3, end_row + 1):
        ledger.row_dimensions[r].height = 22

    if rows:
        table = Table(displayName="FinanceLedgerTable", ref=f"A2:L{end_row}")
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ledger.add_table(table)
        ledger.auto_filter.ref = f"A2:L{end_row}"

    ledger.freeze_panes = "A3"
    ledger.sheet_view.showGridLines = False
    ledger.sheet_view.topLeftCell = "A1"
    ledger.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(OUTPUT)

    out_wb = load_workbook(OUTPUT, data_only=False)
    after_source = snapshot_source(out_wb[SOURCE_SHEET])
    report = {
        "output": str(OUTPUT),
        "source_sheet_business_data_preserved": before_source == after_source,
        "normalized_rows": len(rows),
        "source_sheet": SOURCE_SHEET,
        "ledger_sheet": LEDGER_SHEET,
        "ledger_range": f"A1:L{end_row}",
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
