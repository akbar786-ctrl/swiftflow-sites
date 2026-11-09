"""
Compares a workbook produced by the Excel macro (run on the blank template) with
the workbook produced by build.py: every formula and value on every calculation
page, the Match Day Resource headers/totals, and the Match Day Resource cell
colours (so the Super High block and the risk-first move are identical).

Usage: python3 compare_vba_build.py <macro-built.xlsx> <build.py output.xlsx>
"""
import re
import sys
import warnings
from collections import Counter

import openpyxl

warnings.filterwarnings("ignore")
CALC = ["Demand Sheet", "Cost Sheet", "MD Demand Template", "MD Cost Template", "Fleet", "Cost Calculator",
        "Wider Tournament Template", "Command Team & Planning", "Other Resources", "Baseline Costs",
        "GDP Deflators", "WT Cost Sheet", "WT Demand Sheet"]


def norm(v):
    if isinstance(v, str):
        v = v.replace("_xlfn.", "").replace(" ", "").replace(",FALSE())", ",0)").replace(",FALSE)", ",0)")
        v = re.sub(r"(?<![A-Za-z'])\$?([A-Z]{1,3})\$?(\d+)", r"\1\2", v) if v.startswith("=") else v
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def fill(c):
    return c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None


def main(a, b):
    wa, wb = openpyxl.load_workbook(a), openpyxl.load_workbook(b)
    stats, ex = Counter(), []
    print("sheet order same:", wa.sheetnames == wb.sheetnames, wa.sheetnames)
    print("hidden same:", [s.sheet_state for s in wa] == [s.sheet_state for s in wb])
    for name in CALC:
        sa, sb = wa[name], wb[name]
        rows, cols = max(sa.max_row, sb.max_row), max(sa.max_column, sb.max_column)
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                va, vb = norm(sa.cell(r, c).value), norm(sb.cell(r, c).value)
                if va in ("", None) and vb in ("", None):
                    continue
                stats["cells"] += 1
                same = va == vb or (isinstance(va, (int, float)) and isinstance(vb, (int, float)) and abs(va - vb) < 1e-9) \
                    or (hasattr(va, "date") and hasattr(vb, "date") and va == vb)
                if not same:
                    stats["diff:" + name] += 1
                    if len(ex) < 15:
                        ex.append((name, sa.cell(r, c).coordinate, sa.cell(r, c).value, sb.cell(r, c).value))
    ma, mb = wa["Match Day Resource"], wb["Match Day Resource"]
    for r in range(1, 345):
        for c in range(1, 70):
            va, vb = norm(ma.cell(r, c).value), norm(mb.cell(r, c).value)
            if (va or vb) and va != vb:
                stats["diff:MDR values"] += 1
                if len(ex) < 30:
                    ex.append(("MDR", ma.cell(r, c).coordinate, ma.cell(r, c).value, mb.cell(r, c).value))
            if r >= 3 and 4 <= c <= 63 and fill(ma.cell(r, c)) != fill(mb.cell(r, c)):
                stats["diff:MDR colours"] += 1
                if stats["diff:MDR colours"] <= 5:
                    ex.append(("MDR colour", ma.cell(r, c).coordinate, fill(ma.cell(r, c)), fill(mb.cell(r, c))))
            stats["mdr cells"] += 1
    m3a = sorted(str(m) for m in ma.merged_cells.ranges if m.min_row == 3)
    m3b = sorted(str(m) for m in mb.merged_cells.ranges if m.min_row == 3)
    print("row 3 merges same:", m3a == m3b, m3a)
    print(dict(stats))
    for e in ex:
        print("  ", e)
    return stats


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
