"""
Test pack for the Power BI build: three SYNTHETIC returns (made-up numbers in the forces'
template), a reference file with test fixtures, and the figures Power BI should show
(calculated by pbi_mirror.py, the same logic that reproduced the original workbook exactly).
"""
import datetime as dt
import os
import random
import sys
import warnings
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import make_reference as MR  # noqa: E402
import pbi_mirror as PM  # noqa: E402

TEMPLATE = os.path.join(HERE, "output", "EURO_2028_Resource_Collection_Template.xlsx")
OUT = os.path.join(HERE, "samples")
FORCES = {"Sample Force A": 101, "Sample Force B": 202, "Sample Force C": 303}
FIXTURES = [("A1", dt.date(2028, 6, 9), "High", "Sample Force A"), ("A2", dt.date(2028, 6, 12), "Medium", "Sample Force A"),
            ("A3", dt.date(2028, 6, 15), "Low", "Sample Force A"), ("B1", dt.date(2028, 6, 10), "Medium", "Sample Force B"),
            ("B2", dt.date(2028, 6, 11), "Super High", "Sample Force B"), ("C1", dt.date(2028, 6, 20), "High", "Sample Force C"),
            ("C1", dt.date(2028, 6, 20), "High", "Sample Force B")]
ROLES = ["ARV", "CTSFO", "CBRN - Search", "PO Bronze Commander", "Medic"]


def allowed_days(rule):
    return {"NONMD": ["MD-2", "MD-1", "MD+1"], "MDONLY": ["MD"]}.get(rule, ["MD-2", "MD-1", "MD", "MD+1", "WT"])


def fill(path, seed, ref):
    rnd = random.Random(seed)
    wb = openpyxl.load_workbook(path)
    for sheet, rowmap, colmap in ((PM.MD_SHEET, ref["rowmap_md"], ref["colmap_md"]),
                                  (PM.WT_SHEET, ref["rowmap_wt"], ref["colmap_wt"])):
        ws = wb[sheet]
        for m in rowmap:
            r, lt = m["Excel row"], m["Line type"]
            if lt not in ("PSU config", "PSU count", "MA PSU count") and rnd.random() < 0.6:
                continue
            if m["Rank from column C"] == "Yes":
                ws.cell(r, 3, rnd.choice(["Inspector", "Sergeant", "Constable"]))
            if lt == "Role":
                ws.cell(r, 3, rnd.choice(ROLES))
            if lt == "Additional":
                ws.cell(r, 2, rnd.choice(["Welfare Officer", "Event Liaison"]))
            for c in colmap:
                if c["Day"] not in allowed_days(m["Day rule"]) or rnd.random() < 0.25:
                    continue
                s = MR.RISKS.index(c["Risk"])
                if c["Measure"] == "Number":
                    v = {"PSU config": {"Inspector": 1, "Sergeant": 3, "Constable": 18}.get(m["Rank"], 1),
                         "PSU count": 1 + s + rnd.randint(0, 2), "MA PSU count": rnd.randint(0, 1 + s)}.get(
                        lt, rnd.randint(1, 4) * (s + 1))
                elif c["Measure"] == "Normal ToD":
                    v = rnd.choice([8, 10])
                else:
                    v = rnd.choice([0, 2, 4, 4.5, 6])
                ws.cell(c["Column number"] and r, c["Column number"], v)
    wb.save(path)


def expected(md, wt, cal, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Check Power BI"
    ws["A1"] = ("Expected Power BI results for the SAMPLE returns (synthetic numbers). Build a table visual with the "
                "rows shown and these measures - every figure should match.")
    ws["A1"].font = Font(bold=True)
    r = [3]

    def block(title, head, rows):
        ws.cell(r[0], 1, title).font = Font(bold=True, size=12, color="1F3864")
        r[0] += 1
        for j, h in enumerate(head, 1):
            ws.cell(r[0], j, h).font = Font(bold=True)
        r[0] += 1
        for row in rows:
            for j, v in enumerate(row, 1):
                c = ws.cell(r[0], j, round(v, 2) if isinstance(v, float) else v)
                if isinstance(v, float):
                    c.number_format = "#,##0.00"
                if isinstance(v, dt.date):
                    c.number_format = "dd/mm/yyyy"
            r[0] += 1
        r[0] += 1

    order = {k: i for i, k in enumerate(MR.RISKS)}

    def agg(rows, keys, fields):
        acc = defaultdict(lambda: [0.0] * len(fields))
        for x in rows:
            for i, f in enumerate(fields):
                acc[tuple(x[k] for k in keys)][i] += x[f]
        return [list(k) + v for k, v in sorted(acc.items(), key=lambda kv: [order.get(e, e) for e in kv[0]])]

    block("MD_Fact - one fixture at each risk level. Rows: Force, Risk",
          ["Force", "Risk", "MD Officer Deployments", "MD PSUs Serials", "MD Overtime Cost", "MD Mutual Aid PSU Cost",
           "MD Cost (Original Basis)"],
          [k + [k[4] + k[5]] for k in agg(md, ["Force", "Risk"], ["Officers", "PSUs", "Overtime", "Package"])])
    block("WT_Fact. Rows: Force, Risk",
          ["Force", "Risk", "WT Daily Officers", "WT Officer Days (Tournament)", "WT Overtime Cost (Tournament)",
           "WT Mutual Aid PSU Cost (Tournament)", "WT Cost (Original Basis)"],
          [k + [k[4] + k[5]] for k in agg(wt, ["Force", "Risk"], ["Officers", "TOfficers", "TOvertime", "TPackage"])])
    block("Calendar_Fact. Rows: Force, Calendar date",
          ["Force", "Date", "Calendar Officer Deployments", "Calendar Overtime Cost", "Calendar Mutual Aid PSU Cost",
           "Calendar Cost (Original Basis)"],
          [k + [k[3] + k[4]] for k in agg(cal, ["Force", "Date"], ["Officers", "Overtime", "Package"])])
    t = lambda rows, f: sum(x[f] for x in rows)  # noqa: E731
    block("Cards (no filters)", ["Measure", "Value"], [
        ["Forces Returned", len({x["Force"] for x in md})],
        ["MD Cost (Original Basis)", t(md, "Overtime") + t(md, "Package")],
        ["Calendar Cost (Original Basis)", t(cal, "Overtime") + t(cal, "Package")],
        ["WT Cost (Original Basis)", t(wt, "TOvertime") + t(wt, "TPackage")],
        ["Tournament Cost (Original Basis)", t(cal, "Overtime") + t(cal, "Package") + t(wt, "TOvertime") + t(wt, "TPackage")]])
    for col, w in zip("ABCDEFG", [22, 14, 26, 26, 28, 30, 26]):
        ws.column_dimensions[col].width = w
    wb.save(out)


def main():
    os.makedirs(os.path.join(OUT, "returns"), exist_ok=True)
    ref_path = MR.build(os.path.join(OUT, "EURO_2028_Power_BI_Reference_SAMPLE.xlsx"), FIXTURES)
    ref = PM.load_reference(ref_path)
    paths = []
    for name, seed in FORCES.items():
        p = os.path.join(OUT, "returns", f"{name}.xlsx")
        with open(TEMPLATE, "rb") as a, open(p, "wb") as b:
            b.write(a.read())
        fill(p, seed, ref)
        paths.append(p)
    md, wt, cal, _ = PM.run(paths, ref_path)
    expected(md, wt, cal, os.path.join(OUT, "EXPECTED_Power_BI_results_SAMPLE.xlsx"))
    print(len(md), "MD rows,", len(wt), "WT rows,", len(cal), "calendar rows")


if __name__ == "__main__":
    main()
