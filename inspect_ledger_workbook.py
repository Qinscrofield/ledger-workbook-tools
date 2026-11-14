from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


INPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ledger.xlsx")


def cell_text(value):
    if value is None:
        return ""
    return str(value)


def main() -> None:
    wb = load_workbook(INPUT, data_only=False)
    summary = []
    for ws in wb.worksheets:
        non_empty = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    non_empty.append((cell.row, cell.column, cell.value))
        if non_empty:
            min_row = min(r for r, _, _ in non_empty)
            max_row = max(r for r, _, _ in non_empty)
            min_col = min(c for _, c, _ in non_empty)
            max_col = max(c for _, c, _ in non_empty)
        else:
            min_row = max_row = min_col = max_col = 1

        preview = []
        for r in range(min_row, min(max_row, min_row + 9) + 1):
            preview.append(
                [
                    cell_text(ws.cell(r, c).value)
                    for c in range(min_col, min(max_col, min_col + 11) + 1)
                ]
            )

        header_scores = []
        for r in range(min_row, min(max_row, min_row + 14) + 1):
            vals = [cell_text(ws.cell(r, c).value).strip() for c in range(min_col, max_col + 1)]
            joined = "|".join(vals)
            score = sum(
                1
                for key in ["日期", "金额", "收入", "支出", "收支", "备注", "编号", "对象", "单位", "摘要", "用途"]
                if key in joined
            )
            non_blank = sum(1 for v in vals if v)
            header_scores.append((score, non_blank, r, vals[:12]))
        header_scores.sort(reverse=True)

        summary.append(
            {
                "sheet": ws.title,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "used_range": f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}",
                "non_empty_cells": len(non_empty),
                "merged_ranges": [str(rng) for rng in ws.merged_cells.ranges],
                "likely_header_rows": header_scores[:3],
                "preview": preview,
            }
        )
    print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
