from __future__ import annotations

import json
import sys
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


INPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ledger.xlsx")
OUTPUT_DIR = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("output")
OUTPUT = OUTPUT_DIR / (INPUT.stem + "_规范优化.xlsx")
REPORT = OUTPUT_DIR / "preservation_report.json"

SHEET_NAME = "Sheet1"
HEADER_ROWS = (2, 81)
SUMMARY_ROWS = range(80, 85)
DATA_MIN_ROW = 1
DATA_MAX_ROW = 84
DATA_MIN_COL = 1
DATA_MAX_COL = 52


def snapshot(wb):
    data = {}
    for ws in wb.worksheets:
        rows = []
        for r in range(1, ws.max_row + 1):
            rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        data[ws.title] = {
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "values": rows,
            "merged_ranges": [str(rng) for rng in ws.merged_cells.ranges],
        }
    return data


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def same_snapshot(before, after):
    return before == after


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(INPUT, data_only=False)
    before = snapshot(wb)

    ws = wb[SHEET_NAME]
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = "A2:AZ84"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:2"

    thin = Side(style="thin", color="D9E2EC")
    medium = Side(style="medium", color="8FA7BD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    section_border = Border(left=thin, right=thin, top=medium, bottom=thin)

    title_fill = PatternFill("solid", fgColor="1F4E5F")
    header_fill = PatternFill("solid", fgColor="2F75B5")
    header_fill_2 = PatternFill("solid", fgColor="5B9BD5")
    summary_fill = PatternFill("solid", fgColor="EAF3F8")
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    blank_fill = PatternFill("solid", fgColor="F7FAFC")

    base_font = Font(name="Microsoft YaHei", size=10, color="1F2933")
    header_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    summary_font = Font(name="Microsoft YaHei", size=10, bold=True, color="1F2933")

    date_cols = set()
    remark_cols = set()
    for c in range(DATA_MIN_COL, DATA_MAX_COL + 1):
        header = ws.cell(2, c).value
        if header == "\u65e5\u671f":
            date_cols.add(c)
        if isinstance(header, str) and "\u5907\u6ce8" in header:
            remark_cols.add(c)

    for row in ws.iter_rows(
        min_row=DATA_MIN_ROW,
        max_row=DATA_MAX_ROW,
        min_col=DATA_MIN_COL,
        max_col=DATA_MAX_COL,
    ):
        for cell in row:
            cell.font = copy(base_font)
            cell.border = border
            cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=False)

            if cell.value is None:
                cell.fill = blank_fill
            if cell.column in remark_cols:
                cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
            if is_number(cell.value) or is_formula(cell.value):
                cell.number_format = '#,##0.00;[Red]-#,##0.00;"-"'
                cell.alignment = Alignment(vertical="center", horizontal="right", wrap_text=False)
            if cell.column in date_cols and 3 <= cell.row <= 79:
                cell.number_format = "@"
                cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=False)

    for rng in ws.merged_cells.ranges:
        for row in ws[str(rng)]:
            for cell in row:
                cell.border = border

    ws["B1"].font = title_font
    ws["B1"].fill = title_fill
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    for c in range(2, 39):
        ws.cell(1, c).fill = title_fill
        ws.cell(1, c).border = border

    for r in HEADER_ROWS:
        for c in range(DATA_MIN_COL, DATA_MAX_COL + 1):
            cell = ws.cell(r, c)
            cell.font = copy(header_font)
            cell.fill = header_fill if r == 2 else header_fill_2
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=medium, bottom=medium)

    for r in SUMMARY_ROWS:
        for c in range(DATA_MIN_COL, DATA_MAX_COL + 1):
            cell = ws.cell(r, c)
            cell.fill = total_fill if r in (80, 82, 83, 84) else header_fill_2
            cell.border = section_border if r == 80 else border
            if r == 81:
                cell.font = copy(header_font)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = copy(summary_font)

    for r in range(3, 80):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 28
    for r in SUMMARY_ROWS:
        ws.row_dimensions[r].height = 25

    for c in range(DATA_MIN_COL, DATA_MAX_COL + 1):
        letter = get_column_letter(c)
        header = ws.cell(2, c).value
        width = 12
        if c in date_cols:
            width = 11
        elif c in remark_cols:
            width = 20
        elif isinstance(header, str) and any(k in header for k in ["\u7a0e", "\u4ed8\u6b3e", "\u8f6c\u5165", "\u8f6c\u51fa", "\u5f00\u652f", "\u6750\u6599", "\u65bd\u5de5", "\u540a\u7bee"]):
            width = 15
        elif header is None:
            width = 11
        ws.column_dimensions[letter].width = width

    ws.column_dimensions["H"].width = 16
    ws.column_dimensions["Z"].width = 24
    ws.column_dimensions["AC"].width = 22
    ws.column_dimensions["AF"].width = 18
    ws.column_dimensions["AI"].width = 22
    ws.column_dimensions["AX"].width = 22
    ws.column_dimensions["AY"].width = 10
    ws.column_dimensions["AZ"].width = 18

    for ws_blank in wb.worksheets:
        if ws_blank.title != SHEET_NAME:
            ws_blank.sheet_view.showGridLines = False

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True

    wb.save(OUTPUT)

    out_wb = load_workbook(OUTPUT, data_only=False)
    after = snapshot(out_wb)
    report = {
        "input": str(INPUT),
        "output": str(OUTPUT),
        "preserved_values_formulas_merges_dimensions": same_snapshot(before, after),
        "sheets": list(before.keys()),
        "sheet1_range": "A1:AZ84",
        "unchanged_sheet_count": len(before) == len(after),
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
