"""Helpers to fill a copy of the collection template the way a force would."""
import openpyxl


def _colmap(ws):
    for r in range(1, 15):
        if ws.cell(r, 1).value == "ROWKEY":
            return {ws.cell(r, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
    raise ValueError("key row not found")


def _rowmap(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}


def fill(template, out, cover=None, status=None, values=None, attrs=None):
    """
    cover:  {"cv_Force": "...", ...}
    status: {"MD-PSC": "Complete", ...}
    values: {("MD", "MD-PSC-07|1", "H|MD|NUM"): 1, ...}
    attrs:  {("MD", "MD-ADD-01|1", "RANK"): "Constable", ...}  (RANK/ROLE/SUPPLY/LOCATION)
    """
    wb = openpyxl.load_workbook(template)
    for name, val in (cover or {}).items():
        sheet, ref = next(iter(wb.defined_names[name].destinations))
        wb[sheet][ref.replace("$", "")] = val
    maps = {}
    for coll in ("MD", "WT"):
        ws = wb[f"{coll} Collection"]
        maps[coll] = (ws, _colmap(ws), _rowmap(ws))
    for sec, st in (status or {}).items():
        ws, cm, rm = maps[sec[:2]]
        ws.cell(rm[f"SECTION|{sec}"], cm["LOCATION"], st)
    for (coll, rk, ck), v in (values or {}).items():
        ws, cm, rm = maps[coll]
        ws.cell(rm[rk], cm[ck], v)
    for (coll, rk, ck), v in (attrs or {}).items():
        ws, cm, rm = maps[coll]
        ws.cell(rm[rk], cm[ck], v)
    wb.save(out)
    return out
