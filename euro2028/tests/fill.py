"""Helpers to fill a copy of the collection template the way a force would."""
import os
import sys

import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "build"))
import catalogue as C  # noqa: E402


def sheet_maps(ws):
    """(column map {key: col}, row map {rowkey: row}) using the hidden key row / column."""
    for r in range(1, 15):
        vals = {ws.cell(r, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
        if "ROWKEY" in vals:
            kc = vals["ROWKEY"]
            rows = {ws.cell(x, kc).value: x for x in range(r + 1, ws.max_row + 1) if ws.cell(x, kc).value}
            return vals, rows
    raise ValueError("key row not found")


def fill(template, out, cover=None, status=None, values=None, attrs=None, all_complete=False):
    """
    cover:  {"cv_Force": "...", ...}
    status: {"MD-PSC": "Complete", ...}   (Cover sheet section table)
    values: {("MD", "MD-PSC-07|1", "H|MD|NUM"): 1, ...}
    attrs:  {("MD", "MD-ADD-01|1", "COLB"): "Hotel liaison", ...}   COLA / COLB / COLC
    """
    wb = openpyxl.load_workbook(template)
    for name, val in (cover or {}).items():
        sheet, ref = next(iter(wb.defined_names[name].destinations))
        wb[sheet][ref.replace("$", "")] = val
    ws = wb["Cover"]
    t = ws.tables["tblSectionStatus"]
    rows = list(ws[t.ref])
    hdr = [c.value for c in rows[0]]
    for rw in rows[1:]:
        key = rw[hdr.index("SectionKey")].value
        if all_complete:
            rw[hdr.index("Status")].value = "Complete"
        if status and key in status:
            rw[hdr.index("Status")].value = status[key]
    maps = {coll: (wb[name], *sheet_maps(wb[name])) for coll, name in C.SHEET_NAMES.items()}
    for (coll, rk, ck), v in list((values or {}).items()) + list((attrs or {}).items()):
        ws, cm, rm = maps[coll]
        ws.cell(rm[rk], cm[ck], v)
    wb.save(out)
    return out
