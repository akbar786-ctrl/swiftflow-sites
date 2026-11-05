"""
One-off extraction of the legacy pages that do not depend on the collection
layout (Baseline Costs, GDP Deflators, Command Team & Planning, Other Resources)
from the original workbook, with styles, into legacy_static_pages.json.

Force-entered inputs are blanked so no force data is carried into the template:
  Command Team & Planning  B23:C28 (post start / end months)
  Other Resources          D7:D26, F7:F26 (hours)

Usage: python3 extract_static_pages.py <original workbook.xlsx> <out.json>
"""
import datetime as dt
import json
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")
SHEETS = ["Baseline Costs", "GDP Deflators", "Command Team & Planning", "Other Resources"]
BLANK = {"Command Team & Planning": lambda c, r: c in ("B", "C") and 23 <= r <= 28,
         "Other Resources": lambda c, r: c in ("D", "F") and 7 <= r <= 26}


def colour(c):
    try:
        if c is not None and c.type == "rgb" and isinstance(c.rgb, str) and c.rgb not in ("00000000",):
            return c.rgb
        if c is not None and c.type == "theme":
            return f"theme:{c.theme}:{round(c.tint or 0, 3)}"
    except Exception:
        pass
    return None


def main(src, out):
    wb = openpyxl.load_workbook(src)
    data = {}
    for name in SHEETS:
        ws = wb[name]
        cells = []
        for row in ws.iter_rows():
            for c in row:
                if type(c).__name__ == "MergedCell":
                    continue
                v = c.value
                if v is None and not c.has_style:
                    continue
                col = c.column_letter
                if v is not None and name in BLANK and BLANK[name](col, c.row) and not (isinstance(v, str) and v.startswith("=")):
                    if not isinstance(v, str):
                        v = None
                if isinstance(v, (dt.datetime, dt.date)):
                    v = {"date": v.isoformat()[:10]}
                f = c.font
                cells.append(dict(
                    ref=c.coordinate, v=v, nf=c.number_format,
                    b=bool(f.b), i=bool(f.i), sz=f.sz, fc=colour(f.color),
                    fill=colour(c.fill.fgColor) if c.fill and c.fill.fill_type == "solid" else None,
                    h=c.alignment.horizontal, wrap=bool(c.alignment.wrap_text),
                    bd=bool(c.border and (c.border.left.style or c.border.top.style))))
        data[name] = dict(
            cells=[x for x in cells if x["v"] is not None or x["fill"] or x["bd"]],
            widths={k: d.width for k, d in ws.column_dimensions.items() if d.width},
            merges=[str(m) for m in ws.merged_cells.ranges],
            state=ws.sheet_state)
    json.dump(data, open(out, "w"), indent=0, default=str)
    print({k: len(v["cells"]) for k, v in data.items()})


if __name__ == "__main__":
    main(*sys.argv[1:3])
