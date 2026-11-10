"""
The Power BI model's logic in Python, step for step (powerbi/*.pq), so it can be tested
against the original collection workbook and give expected figures for checking Power BI.

Steps (same names as the Power Query queries):
  Cells     read MD_Data / WT_Data (A1-anchored grids) -> one row per Excel row x numeric column
  Lines     + row map + column map -> per force / row / risk / day: Number, Normal ToD, Extra Hours
  MD_Fact   officer lines; PSU lines = PSU count x officers per PSU (configuration rows);
            mutual aid PSU = count x package; overtime cost = officers x extra hours x rate
  WT_Fact   the same for the Wider Tournament sheet, x days at each risk level
  Calendar_Fact  MD_Fact placed on dates from the fixture list (original one-day-per-date rule)
"""
import datetime as dt
import os
from collections import defaultdict

import openpyxl

MD_SHEET, WT_SHEET = "Match Day Resource", "Wider Tournament Resource"


def tbl(wb, name):
    for ws in wb:
        if name in ws.tables:
            rng = ws[ws.tables[name].ref]
            head = [c.value for c in rng[0]]
            return [dict(zip(head, [c.value for c in r])) for r in rng[1:]]
    raise KeyError(name)


def load_reference(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ref = dict(rowmap_md=tbl(wb, "tblRowMapMD"), rowmap_wt=tbl(wb, "tblRowMapWT"), colmap_md=tbl(wb, "tblColMapMD"),
               colmap_wt=tbl(wb, "tblColMapWT"))
    ref["rates"] = {(r["Supplied by"], r["Rank"]): r["Hourly rate"] for r in tbl(wb, "tblRates") if r["Rank"]}
    ref["package"] = {r["Key"]: r["Value"] for r in tbl(wb, "tblSettings")}["MA_PSU_PACKAGE"]
    ref["wt_days"] = {r["Risk"]: r["Days"] for r in tbl(wb, "tblWTDays") if r["Risk"]}
    ref["fixtures"] = [(str(r["Fixture"] or ""), r["Date"].date() if isinstance(r["Date"], dt.datetime) else r["Date"],
                        r["Risk"], r["Force"]) for r in tbl(wb, "tblFixtures") if r["Date"] and r["Risk"] and r["Force"]]
    return ref


def norm_rank(x):
    x = (str(x).replace("\xa0", " ").strip() if x is not None else "")
    return {"Sergeants": "Sergeant", "Constables": "Constable", "Supt": "Superintendent"}.get(x, x)


def lines(path, sheet, rowmap, colmap, force):
    """Cells + Lines: per mapped row and risk/day -> Number, Normal ToD, Extra Hours (plus column B/C text)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    cols = {c["Column number"]: c for c in colmap}
    out = []
    for m in rowmap:
        r = m["Excel row"]
        b, c = ws.cell(r, 2).value, ws.cell(r, 3).value
        acc = defaultdict(dict)
        for cn, cm in cols.items():
            v = ws.cell(r, cn).value
            if isinstance(v, str):
                try:
                    v = float(v)
                except ValueError:
                    v = None
            if v in (None, 0) or isinstance(v, bool):
                continue
            acc[(cm["Risk"], cm["Day"])][cm["Measure"]] = float(v)
        for (risk, day), meas in acc.items():
            out.append(dict(Force=force, Row=r, Map=m, B=b, C=c, Risk=risk, Day=day,
                            Number=meas.get("Number", 0.0), ToD=meas.get("Normal ToD", 0.0),
                            Extra=meas.get("Extra Hours", 0.0)))
    return out


def fact(lns, ref, wt=False):
    rates, package = ref["rates"], ref["package"]
    out = []
    by_group = defaultdict(list)           # PSU configuration lines by (force, group, risk, day)
    for x in lns:
        if x["Map"]["Line type"] == "PSU config":
            by_group[(x["Force"], x["Map"]["PSU group"], x["Risk"], x["Day"])].append(x)
    for x in lns:
        m = x["Map"]
        lt = m["Line type"]
        base = dict(Force=x["Force"], Row=x["Row"], Section=m["Section"], Resource=m["Resource"],
                    Supply=m["Supplied by"] or "", Risk=x["Risk"], Day=x["Day"], DayRule=m["Day rule"])
        if lt == "PSU config":
            continue
        if lt in ("PSU count", "MA PSU count"):
            unit_type = str(x["C"] or "").strip()
            out.append(dict(base, LineType="PSU units", Rank="PSUs / Serials", Role=unit_type, Officers=0.0,
                            PSUs=x["Number"], ExtraTotal=0.0, ToDTotal=0.0, Overtime=0.0, NormalCost=0.0,
                            Package=x["Number"] * package if lt == "MA PSU count" else 0.0))
            if lt == "PSU count":
                for cfg in by_group.get((x["Force"], m["PSU group"], x["Risk"], x["Day"]), []):
                    cm = cfg["Map"]
                    officers = x["Number"] * cfg["Number"]
                    rate = rates.get((cm["Rate supplied by"], cm["Rate rank"]), 0.0) or 0.0
                    out.append(dict(base, LineType="PSU officers", Rank=cm["Rank"], Role=unit_type, Officers=officers,
                                    PSUs=0.0, ExtraTotal=officers * cfg["Extra"], ToDTotal=officers * cfg["ToD"],
                                    Overtime=officers * cfg["Extra"] * rate, NormalCost=officers * cfg["ToD"] * rate,
                                    Package=0.0))
            continue
        rank = norm_rank(x["C"]) if m["Rank from column C"] == "Yes" else (m["Rank"] or "")
        role = str(x["C"] or "").strip() if m["Role from column C"] == "Yes" else ""
        resource = m["Resource"]
        if lt == "Additional" and x["B"]:
            resource = str(x["B"]).strip()
        rate_rank = m["Rate rank"] or rank
        rate = rates.get((m["Rate supplied by"], rate_rank), 0.0) or 0.0
        out.append(dict(base, Resource=resource, LineType=lt, Rank=rank, Role=role, Officers=x["Number"], PSUs=0.0,
                        ExtraTotal=x["Number"] * x["Extra"], ToDTotal=x["Number"] * x["ToD"],
                        Overtime=x["Number"] * x["Extra"] * rate, NormalCost=x["Number"] * x["ToD"] * rate,
                        Package=0.0))
    if wt:
        for f in out:
            d = ref["wt_days"].get(f["Risk"], 0)
            f.update(Days=d, TOfficers=f["Officers"] * d, TOvertime=f["Overtime"] * d, TPackage=f["Package"] * d,
                     TNormal=f["NormalCost"] * d)
    return out


PRI = {"STD": {"MD": 1, "MD-1": 2, "MD-2": 3, "MD+1": 4}, "NONMD": {"MD-1": 1, "MD-2": 2, "MD+1": 3},
       "MDONLY": {"MD": 1}}
OFFSET = {"MD-2": -2, "MD-1": -1, "MD": 0, "MD+1": 1}


def calendar(md_fact, fixtures):
    present = defaultdict(set)
    for _, date, risk, force in fixtures:
        for day, o in OFFSET.items():
            present[(force, date + dt.timedelta(o), risk)].add(day)
    pick = {}
    for (force, date, risk), days in present.items():
        for rule, pr in PRI.items():
            cand = [d for d in days if d in pr]
            if cand:
                pick.setdefault((force, risk, min(cand, key=lambda d: pr[d]), rule), []).append(date)
    return [dict(f, Date=d) for f in md_fact for d in pick.get((f["Force"], f["Risk"], f["Day"], f["DayRule"]), [])]


def run(paths, ref_path, forces=None):
    ref = load_reference(ref_path)
    md_l, wt_l = [], []
    for i, p in enumerate(paths):
        force = (forces or {}).get(p) or os.path.splitext(os.path.basename(p))[0]
        md_l += lines(p, MD_SHEET, ref["rowmap_md"], ref["colmap_md"], force)
        wt_l += lines(p, WT_SHEET, ref["rowmap_wt"], ref["colmap_wt"], force)
    md, wt = fact(md_l, ref), fact(wt_l, ref, wt=True)
    return md, wt, calendar(md, ref["fixtures"]), ref
