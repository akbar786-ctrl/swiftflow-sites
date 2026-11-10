"""
Proves the Power BI logic reproduces the original collection workbook.

Input: a workbook in the new template layout holding a force's figures AND the original-logic
calculation pages (made by euro2028/v1_lift_shift/validate_against_original.py, which itself
matches the original workbook to 0 differences), recalculated.

The Power BI logic (pbi_mirror.py) reads ONLY Match Day Resource and Wider Tournament Resource
plus the reference tables (fixtures taken from the Demand Sheet markers) and must give the same
  * officers per date (MD Demand Template TOTAL row; mutual aid PSUs count as units there)
  * cost per date (MD Cost Template TOTAL row)
  * Wider Tournament total (Wider Tournament Template TOTAL column).

Usage: python3 validate_against_original.py <filled + recalculated workbook.xlsx> <work dir>
"""
import datetime as dt
import os
import sys
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import make_reference as MR  # noqa: E402
import pbi_mirror as PM  # noqa: E402

FORCE = "Test Force"
START, DAYS_N = dt.date(2028, 6, 7), 34


def main(filled, work):
    os.makedirs(work, exist_ok=True)
    wb = openpyxl.load_workbook(filled, data_only=True)
    dem = wb["Demand Sheet"]
    fixtures = []
    for k in range(DAYS_N):
        for s in range(4):
            v = dem.cell(4, 4 + 5 * k + s).value
            if v not in (None, ""):
                fixtures.append((str(v), START + dt.timedelta(k), MR.RISKS[s], FORCE))
    ref = MR.build(os.path.join(work, "reference.xlsx"), fixtures)
    md, wt, cal, _ = PM.run([filled], ref, {filled: FORCE})
    # original results
    res = {}
    for name in ("MD Demand Template", "MD Cost Template"):
        ws = wb[name]
        tr = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 2).value == "TOTAL")
        res[name] = {START + dt.timedelta(k): ws.cell(tr, 4 + k).value or 0 for k in range(DAYS_N)}
    wtt = wb["Wider Tournament Template"]
    tr = next(r for r in range(1, wtt.max_row + 1) if wtt.cell(r, 2).value == "TOTAL")
    orig_wt = wtt.cell(tr, 12).value or 0
    # Power BI logic results
    off, cost = defaultdict(float), defaultdict(float)
    for f in cal:
        off[f["Date"]] += f["Officers"] + (f["PSUs"] if f["Supply"] == "Mutual aid" else 0)
        cost[f["Date"]] += f["Overtime"] + f["Package"]
    diffs = 0
    print(f"{len(fixtures)} fixtures from the Demand Sheet; {len(md)} MD fact rows, {len(cal)} calendar rows")
    print(f"{'date':12}{'officers orig':>15}{'Power BI':>12}{'cost orig':>16}{'Power BI':>16}")
    for d in sorted(res["MD Demand Template"]):
        a, b = res["MD Demand Template"][d], off.get(d, 0)
        c, e = res["MD Cost Template"][d], cost.get(d, 0)
        bad = abs(a - b) > 1e-6 or abs(c - e) > 0.01
        diffs += bad
        if a or b or c or e:
            print(f"{d.isoformat():12}{a:>15,.0f}{b:>12,.0f}{c:>16,.2f}{e:>16,.2f}{'  <-- DIFF' if bad else ''}")
    pbi_wt = sum(f["TOvertime"] + f["TPackage"] for f in wt)
    wt_bad = abs(orig_wt - pbi_wt) > 0.01
    print(f"Wider Tournament total: original {orig_wt:,.2f}  Power BI logic {pbi_wt:,.2f}{'  <-- DIFF' if wt_bad else ''}")
    tot_o = sum(res["MD Cost Template"].values())
    print(f"Match Day cost, all dates: original {tot_o:,.2f}  Power BI logic {sum(cost.values()):,.2f}")
    print("RESULT:", "0 differences" if not diffs and not wt_bad else f"{diffs + wt_bad} differences")
    return diffs + wt_bad


if __name__ == "__main__":
    sys.exit(main(sys.argv[1], sys.argv[2]))
