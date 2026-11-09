"""
Tests the generated Office Script without Excel:
  1. type-checks it with strict TypeScript against the ExcelScript API members it uses
     (officescript_test/excelscript_mock.ts), no 'any' allowed;
  2. runs it on a model of the BLANK v1 template (values, formulas, fills, merges);
  3. compares every value/formula on the calculation pages, and Match Day Resource values,
     fills and row-3 merges, with the workbook made by build.py.

Usage: python3 test_officescript.py <script.ts/.txt> <build.py output.xlsx> <work dir>
"""
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import warnings

import openpyxl

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
CALC = ["Demand Sheet", "Cost Sheet", "MD Demand Template", "MD Cost Template", "Fleet", "Cost Calculator",
        "Wider Tournament Template", "Command Team & Planning", "Other Resources", "Baseline Costs",
        "GDP Deflators", "WT Cost Sheet", "WT Demand Sheet"]


def fid(c):
    if not (c.fill and c.fill.fill_type):
        return None
    rgb = c.fill.fgColor.rgb
    return rgb if isinstance(rgb, str) else f"theme{c.fill.fgColor.theme}:{c.fill.fgColor.tint}"


def val(v):
    if isinstance(v, dt.datetime):
        return (v - dt.datetime(1899, 12, 30)).days
    if isinstance(v, str) and v.startswith("="):
        v = v.replace("_xlfn.", "").replace(" ", "")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def main(script, built, work):
    os.makedirs(work, exist_ok=True)
    # 1. template model
    wb = openpyxl.load_workbook(SOURCE)
    model = {}
    for ws in wb:
        cells = []
        if ws.title == "Match Day Resource":
            for row in ws.iter_rows(min_row=1, max_row=400, max_col=80):
                for c in row:
                    v = c.value
                    f = v if isinstance(v, str) and v.startswith("=") else None
                    if f:
                        v = None
                    if v is not None or f or fid(c):
                        cells.append([c.row, c.column, v, f, fid(c)])
        model[ws.title] = cells
        model["__meta__" + ws.title] = [[0, 0, ws.sheet_state, None, None]] + \
            [[0, 1, str(m), None, None] for m in ws.merged_cells.ranges]
    json.dump(model, open(os.path.join(work, "model.json"), "w"))
    # 2. compile + run
    ts = os.path.join(work, "script.ts")
    shutil.copy(script, ts)
    text = open(ts).read()
    assert not re.search(r"[:<(,]\s*any\b", text), "explicit 'any' is not allowed in Office Scripts"
    out_js = os.path.join(work, "out.js")
    cmd = ["tsc", "--ignoreDeprecations", "6.0", "--strict", "--noImplicitAny", "--target", "es2019", "--lib", "es2019", "--outFile", out_js,
           os.path.join(HERE, "officescript_test", "excelscript_mock.ts"), ts,
           os.path.join(HERE, "officescript_test", "runner.ts")]
    r = subprocess.run(cmd, capture_output=True, text=True)
    print("tsc:", "OK" if r.returncode == 0 else r.stdout[:3000])
    if r.returncode:
        return
    r = subprocess.run(["node", out_js, os.path.join(work, "model.json"), os.path.join(work, "result.json")],
                       capture_output=True, text=True)
    if r.returncode:
        print(r.stderr[:3000])
        return
    print("script ran:", r.stdout.strip())
    res = json.load(open(os.path.join(work, "result.json")))
    # 3. compare
    bw = openpyxl.load_workbook(built)
    stats, ex = {}, []

    def bump(k, e=None):
        stats[k] = stats.get(k, 0) + 1
        if e and len(ex) < 20:
            ex.append(e)
    print("sheet order same:", list(res) == bw.sheetnames)
    print("hidden same:", [res[n]["visibility"] == "Hidden" for n in res] == [bw[n].sheet_state == "hidden" for n in bw.sheetnames])
    for name in CALC:
        got = {(r, c): val(f if f is not None else v) for r, c, v, f, _ in res[name]["cells"] if v is not None or f is not None}
        exp = {}
        for row in bw[name].iter_rows():
            for c in row:
                if c.value not in (None, ""):
                    exp[(c.row, c.column)] = val(c.value)
        for k in set(got) | set(exp):
            a, b = got.get(k), exp.get(k)
            if isinstance(a, str) and a.startswith("="):
                a = a.replace(" ", "")
            if a != b and not (isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) < 1e-9):
                bump("diff:" + name, (name, k, a, b))
            bump("cells")
    m = {(r, c): (v, f, fill) for r, c, v, f, fill in res["Match Day Resource"]["cells"]}
    ws = bw["Match Day Resource"]
    for r in range(1, 345):
        for c in range(1, 70):
            v, f, fill = m.get((r, c), (None, None, None))
            a, b = val(f if f else v), val(ws.cell(r, c).value)
            if (a or b) and a != b:
                bump("diff:MDR values", ("MDR", ws.cell(r, c).coordinate, a, b))
            hidden_in_merge = r == 3 and (c - 4) % 12 != 0           # covered by a merged banner: not visible
            if r >= 3 and 4 <= c <= 63 and not hidden_in_merge and fill != fid(ws.cell(r, c)):
                bump("diff:MDR fills", ("MDR fill", ws.cell(r, c).coordinate, fill, fid(ws.cell(r, c))))
            bump("mdr cells")
    m3 = sorted(x for x in res["Match Day Resource"]["merges"] if re.match(r"^[A-Z]+3:", x))
    b3 = sorted(str(x) for x in ws.merged_cells.ranges if x.min_row == 3)
    print("row 3 merges same:", m3 == b3, m3)
    print(stats)
    for e in ex:
        print("  ", e)


if __name__ == "__main__":
    main(*sys.argv[1:4])
