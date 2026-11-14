from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.views import Selection


INPUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("ledger.xlsx")
OUTPUT = (
    Path(sys.argv[2])
    if len(sys.argv) > 2
    else INPUT.with_name(INPUT.stem + "_WPS兼容版.xlsx")
)


def snapshot(wb):
    data = {}
    for ws in wb.worksheets:
        values = []
        for r in range(1, ws.max_row + 1):
            values.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        data[ws.title] = {
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "values": values,
            "merged_ranges": [str(rng) for rng in ws.merged_cells.ranges],
        }
    return data


def main() -> None:
    wb = load_workbook(INPUT, data_only=False)
    before = snapshot(wb)

    for ws in wb.worksheets:
        ws.freeze_panes = None
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.pane = None
        ws.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]

    wb.active = 0
    wb.save(OUTPUT)

    out = load_workbook(OUTPUT, data_only=False)
    after = snapshot(out)
    ws = out["Sheet1"]
    report = {
        "output": str(OUTPUT),
        "data_and_merges_preserved": before == after,
        "freeze_panes": ws.freeze_panes,
        "top_left_cell": ws.sheet_view.topLeftCell,
        "selection": [(s.activeCell, s.sqref, s.pane) for s in ws.sheet_view.selection],
    }
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
