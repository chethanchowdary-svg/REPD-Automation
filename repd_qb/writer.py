"""One workbook per run. Upload sheets carry only upload columns; changed cells are light yellow,
rows matched by GPT (Ref ID added) are on top with ID/name cells light green."""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HDR_FILL = PatternFill("solid", start_color="1F3864")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name="Arial", size=10)
CHANGED_FILL = PatternFill("solid", start_color="FFF2CC")
GPT_FILL = PatternFill("solid", start_color="C6EFCE")
WRAP = {"Analyst_comments", "Detail", "Reason"}
GPT_COLS = ("Project ID", "Project_ID", "Project_name", "State_data_plant_ID")


def _is_gpt(r):
    return r.get("_source") == "gpt"


def sheet(wb, title, columns, rows, widths=None):
    ws = wb.create_sheet(title[:31])
    ws.append(columns)
    for c in ws[1]:
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
    rows = sorted(rows, key=lambda r: not _is_gpt(r))
    fills = []
    for n, r in enumerate(rows, start=2):
        ws.append([r.get(c, "") for c in columns])
        for col in r.get("_changed", ()):
            if col in columns:
                fills.append((n, columns.index(col) + 1, CHANGED_FILL))
        if _is_gpt(r):
            fills += [(n, columns.index(c) + 1, GPT_FILL) for c in GPT_COLS if c in columns]
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = BODY_FONT
            c.alignment = Alignment(wrap_text=columns[c.column - 1] in WRAP, vertical="top")
    for rr, cc, f in fills:
        ws.cell(rr, cc).fill = f
    for i, col in enumerate(columns, 1):
        w = (widths or {}).get(col) or (60 if col in WRAP else min(max(len(str(col)) + 2, 12), 36))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def write(path, upd: dict, new_rows, new_cols, id_matches, stats):
    cols = upd["columns"]
    wb = Workbook()
    wb.remove(wb.active)
    sheet(wb, "Lifecycle Capacity", cols, upd["lifecycle_capacity"])
    sheet(wb, "Site Area Changes", cols, upd["site_area"])
    sheet(wb, "Other Changes", cols, upd["other"])
    sheet(wb, "Lead QC Update", cols, upd["qc"])
    sheet(wb, "New Projects", new_cols, new_rows)
    if id_matches:
        sheet(wb, "ID Matches (GPT)", list(id_matches[0].keys()), id_matches,
              widths={"Reason": 70, "REPD name": 32, "DB project name": 32})
    sheet(wb, "Full DB", cols, upd["full"])
    log = upd["log"]
    sheet(wb, "Change Log", list(log.columns) if len(log) else ["Project name", "State data plant ID",
                                                                 "Action", "Detail"],
          log.to_dict("records") if len(log) else [], widths={"Detail": 90})
    ws = wb.create_sheet("Run Summary")
    ws.append(["Metric", "Value"])
    for k, v in stats.items():
        ws.append([k, v])
    for c in ws[1]:
        c.fill, c.font = HDR_FILL, HDR_FONT
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 44, 60
    wb.save(path)
