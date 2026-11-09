"""
Python version of the Power BI calculations (powerbi/*.pq), step for step, used to
produce the figures Power BI should show for the sample returns, so the Power BI
build can be checked number by number.

Usage: python3 reference_calc.py <returns folder> <reference data.xlsx> <expected results.xlsx>
"""
import datetime as dt
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font

UNITS = "Number of PSUs / Serials"
KEY = ["Section", "Resource", "Supplied by", "Role / Skill", "Rank"]


def table_rows(ws, name):
    t = ws.tables[name]
    rng = ws[t.ref]
    head = [c.value for c in rng[0]]
    return [dict(zip(head, [c.value for c in row])) for row in rng[1:]]


def find_table(wb, name):
    for ws in wb:
        if name in ws.tables:
            return table_rows(ws, name)
    return None


def form_long(path, table):
    """fnFormLong"""
    wb = openpyxl.load_workbook(path, data_only=True)
    det = find_table(wb, "tblDetails")
    force = (str(det[0].get("Force") or "").strip() if det else "") or os.path.splitext(os.path.basename(path))[0]
    rows = find_table(wb, table)
    if rows is None:
        raise ValueError("not a form")
    out = []
    for i, r in enumerate(rows, 1):
        for k, v in r.items():
            if k is None or " | " not in k or v is None:
                continue
            parts = k.split(" | ")
            try:
                val = float(v)
            except (TypeError, ValueError):
                val = None
            out.append(dict(Force=force, File=os.path.basename(path), Row=i, **{c: r[c] for c in KEY},
                            Risk=parts[0], Day=parts[1] if len(parts) == 3 else "WT", Measure=parts[-1], Value=val))
    return out


def fact(long_rows, rates, overrides, package, wt_days=None):
    """MD_Fact / WT_Fact"""
    lines = defaultdict(dict)
    for r in long_rows:
        if not r["Value"]:
            continue
        key = (r["Force"], r["File"], r["Row"], *[r[c] for c in KEY], r["Risk"], r["Day"])
        lines[key][r["Measure"]] = lines[key].get(r["Measure"], 0) + r["Value"]
    out = []
    for key, m in lines.items():
        force, file, row, section, resource, supply, role, rank, risk, day = key
        units_line = rank == UNITS
        officers = 0 if units_line else m.get("Officers", 0)
        psus = m.get("Officers", 0) if units_line else 0
        nh, xh = m.get("Normal hrs", 0), m.get("Extra hrs", 0)
        rate = overrides.get((resource, supply, rank), rates.get((supply, rank)))
        f = dict(Force=force, Row=row, Section=section, Resource=resource, Supply=supply, Role=role, Rank=rank,
                 Risk=risk, Day=day, Officers=officers, PSUs=psus, Rate=rate,
                 Overtime=officers * xh * (rate or 0), Normal=officers * nh * (rate or 0),
                 Package=psus * package if units_line and supply == "Mutual aid" else 0,
                 RowType="NON_MD" if "(Non-Matchday)" in (resource or "") else "STD")
        if wt_days is not None:
            d = wt_days.get(risk, 0)
            f.update(Days=d, TOfficers=officers * d, TOvertime=f["Overtime"] * d, TPackage=f["Package"] * d,
                     TNormal=f["Normal"] * d)
        out.append(f)
    return out


def calendar(md_fact, fixtures):
    """Calendar_Fact"""
    off = {"MD-2": (-2, 3, 2), "MD-1": (-1, 2, 1), "MD": (0, 1, None), "MD+1": (1, 4, 3)}
    groups = defaultdict(list)
    for fx, date, risk, force in fixtures:
        for day, (o, p, pn) in off.items():
            groups[(force, date + dt.timedelta(o), risk)].append((day, p, pn, fx))
    lookup = {}
    for (force, date, risk), items in groups.items():
        std = min(items, key=lambda x: x[1])[0]
        non = [x for x in items if x[2] is not None]
        lookup.setdefault((force, risk, std, "STD"), []).append(date)
        if non:
            lookup.setdefault((force, risk, min(non, key=lambda x: x[2])[0], "NON_MD"), []).append(date)
    out = []
    for f in md_fact:
        for date in lookup.get((f["Force"], f["Risk"], f["Day"], f["RowType"]), []):
            out.append(dict(f, Date=date))
    return out


def load_reference(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rates = {(r["Supplied by"], r["Rank"]): r["Hourly rate"] for r in find_table(wb, "tblRates") if r["Rank"]}
    ovr = {(r["Resource"], r["Supplied by"], r["Rank"]): r["Hourly rate"]
           for r in find_table(wb, "tblRateOverrides") if r["Resource"]}
    package = {r["Key"]: r["Value"] for r in find_table(wb, "tblSettings")}["MA_PSU_PACKAGE"]
    days = {r["Risk"]: r["Days"] for r in find_table(wb, "tblWTDays") if r["Risk"]}
    fixtures = []
    for r in find_table(wb, "tblFixtures"):
        if r["Date"] and r["Risk"] and r["Force"]:
            d = r["Date"].date() if isinstance(r["Date"], dt.datetime) else r["Date"]
            fixtures.append((str(r["Fixture"] or ""), d, str(r["Risk"]).strip(), str(r["Force"]).strip()))
    return rates, ovr, package, days, fixtures


def run(folder, ref):
    rates, ovr, package, days, fixtures = load_reference(ref)
    md_long, wt_long = [], []
    for name in sorted(os.listdir(folder)):
        if not name.lower().endswith(".xlsx") or name.startswith("~$"):
            continue
        try:
            md_long += form_long(os.path.join(folder, name), "tblMatchDay")
            wt_long += form_long(os.path.join(folder, name), "tblWiderTournament")
        except ValueError:
            continue
    md = fact(md_long, rates, ovr, package)
    wt = fact(wt_long, rates, ovr, package, days)
    cal = calendar(md, fixtures)
    return md, wt, cal


def write_expected(md, wt, cal, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Check Power BI"
    ws["A1"] = "Expected Power BI results for the SAMPLE returns (synthetic data). Put these measures in a table " \
               "visual with the rows shown and compare."
    ws["A1"].font = Font(bold=True)
    r = 3

    def block(title, head, rows):
        nonlocal r
        ws.cell(r, 1, title).font = Font(bold=True, color="1F3864", size=12)
        r += 1
        for j, h in enumerate(head, 1):
            ws.cell(r, j, h).font = Font(bold=True)
        r += 1
        for row in rows:
            for j, v in enumerate(row, 1):
                c = ws.cell(r, j, v)
                if isinstance(v, float):
                    c.number_format = "#,##0.00"
                if isinstance(v, dt.date):
                    c.number_format = "dd/mm/yyyy"
            r += 1
        r += 1

    def agg(rows, keys, fields):
        acc = defaultdict(lambda: [0.0] * len(fields))
        for x in rows:
            k = tuple(x[k] for k in keys)
            for i, f in enumerate(fields):
                acc[k][i] += x[f]
        order = {"Low": 0, "Medium": 1, "High": 2, "Super High": 3}
        return [list(k) + [round(v, 2) for v in vals]
                for k, vals in sorted(acc.items(), key=lambda kv: [order.get(x, x) if not isinstance(x, dt.date) else x
                                                                   for x in kv[0]])]

    block("Match Day - one fixture at each risk level (MD_Fact)  -  rows: Force, Risk",
          ["Force", "Risk", "MD Officer Deployments", "MD PSUs Serials", "MD Overtime Cost",
           "MD Mutual Aid PSU Cost", "MD Normal Duty Cost", "MD Cost (Original Basis)"],
          [k + [round(k[4] + k[5], 2)] for k in agg(md, ["Force", "Risk"], ["Officers", "PSUs", "Overtime", "Package", "Normal"])])
    block("Wider Tournament (WT_Fact)  -  rows: Force, Risk",
          ["Force", "Risk", "WT Daily Officers", "WT Officer Days (Tournament)", "WT Overtime Cost (Tournament)",
           "WT Mutual Aid PSU Cost (Tournament)", "WT Normal Duty Cost (Tournament)", "WT Cost (Original Basis)"],
          [k + [round(k[4] + k[5], 2)] for k in agg(wt, ["Force", "Risk"], ["Officers", "TOfficers", "TOvertime", "TPackage",
                                                                  "TNormal"])])
    block("Fixture calendar (Calendar_Fact)  -  rows: Force, Date",
          ["Force", "Date", "Calendar Officer Deployments", "Calendar Overtime Cost", "Calendar Mutual Aid PSU Cost",
           "Calendar Cost (Original Basis)"],
          [k + [round(k[3] + k[4], 2)] for k in agg(cal, ["Force", "Date"], ["Officers", "Overtime", "Package"])])
    tot = lambda rows, f: round(sum(x[f] for x in rows), 2)  # noqa: E731
    block("Grand totals (cards)", ["Measure", "Value"], [
        ["MD Officer Deployments", tot(md, "Officers")],
        ["MD Cost (Original Basis)", round(tot(md, "Overtime") + tot(md, "Package"), 2)],
        ["WT Cost (Original Basis)", round(tot(wt, "TOvertime") + tot(wt, "TPackage"), 2)],
        ["Calendar Cost (Original Basis)", round(tot(cal, "Overtime") + tot(cal, "Package"), 2)],
        ["Tournament Cost (Original Basis)", round(tot(cal, "Overtime") + tot(cal, "Package")
                                                   + tot(wt, "TOvertime") + tot(wt, "TPackage"), 2)]])
    for col, w in zip("ABCDEFGH", [20, 14, 26, 24, 26, 30, 30, 26]):
        ws.column_dimensions[col].width = w
    wb.save(out)


if __name__ == "__main__":
    md, wt, cal = run(sys.argv[1], sys.argv[2])
    write_expected(md, wt, cal, sys.argv[3])
    print(len(md), "MD fact rows,", len(wt), "WT fact rows,", len(cal), "calendar rows ->", sys.argv[3])
