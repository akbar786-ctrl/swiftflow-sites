"""
Builds a SYNTHETIC completed submission (collection template) and the same
inputs loaded into the backup calculation workbook, for Power BI development
and reconciliation testing.

All figures are invented test values - they are NOT forecasts and must not be
reported. The force name is deliberately not a real force. The return contains
deliberate data-quality problems so the DQ report can be exercised.
"""
import datetime as dt
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [HERE, os.path.join(HERE, "..", "tests")]
import build_calc_workbook  # noqa: E402
import build_template  # noqa: E402
import catalogue as C  # noqa: E402
from fill import fill  # noqa: E402

RISK_SCALE = {"L": 1.0, "M": 1.3, "H": 1.7, "SH": 2.2}
DAY_SCALE = {"MD-2": 0.3, "MD-1": 0.6, "MD": 1.0, "MD+1": 0.4}
FORCE = "SYNTHETIC TEST FORCE"
# Synthetic fixture pattern for the backup workbook (dates in the 2028 window; NOT real fixtures)
SYNTHETIC_FIXTURES = [(dt.date(2028, 6, 12), "H"), (dt.date(2028, 6, 15), "M"), (dt.date(2028, 6, 18), "H"),
                      (dt.date(2028, 6, 21), "M"), (dt.date(2028, 6, 26), "SH"), (dt.date(2028, 7, 1), "L")]


def _days(res):
    if res["collection"] == "WT":
        return ["WT"]
    return [d for d in DAY_SCALE if C.applicable(res, "NUM", d)]


def sample_inputs(with_errors=True):
    v, a = {}, {}
    status = {f"{c}-{s}": "Complete" for (c, s) in C.SECTIONS}
    for res in C.RESOURCES:
        if res["row_type"] != "PSU_CONFIG":
            continue
        comp = {"INSP": 1, "SGT": 3, "PC": 18} if "Level 2" in res["name"] else {"INSP": 0, "SGT": 1, "PC": 7}
        for ln in C.lines(res):
            rk = C.row_key(res["resource_id"], ln["line"])
            for risk in RISK_SCALE:
                for d in _days(res):
                    v[(res["collection"], rk, f"{risk}|{d}|NUM")] = comp[ln["rank"]]
                    v[(res["collection"], rk, f"{risk}|{d}|TOD")] = 10
                    v[(res["collection"], rk, f"{risk}|{d}|XH")] = 4 if d == "MD" else 2
    for rid, base in [("MD-POP-01", 2), ("MD-POP-02", 1), ("MD-POP-03", 4), ("MD-POP-04", 2), ("MD-POP-05", 6),
                      ("MD-POP-06", 3), ("MD-POP-07", 2), ("MD-POP-08", 1),
                      ("WT-POP-01", 2), ("WT-POP-02", 1), ("WT-POP-03", 3), ("WT-POP-04", 1)]:
        res = C.RESOURCE_BY_ID[rid]
        for ln in C.lines(res):
            for risk, rs in RISK_SCALE.items():
                for d in _days(res):
                    v[(res["collection"], C.row_key(rid, ln["line"]), f"{risk}|{d}|NUM")] = \
                        round(base * rs * DAY_SCALE.get(d, 1) / ln["line"])
    picks = {"MD-EGT-01": 4, "MD-EGT-02": 2, "MD-MNT-03": 8, "MD-DOG-05": 4, "MD-FBP-01": 6, "MD-FBP-02": 4,
             "MD-RDS-01": 10, "MD-ARM-03": 6, "MD-ARM-05": 4, "MD-SRC-01": 3, "MD-UAS-03": 2, "MD-LFR-03": 2,
             "MD-CBR-01": 3, "MD-CMD-01": 1, "MD-CTL-01": 5, "MD-CUS-01": 4, "MD-CRM-01": 3, "MD-CRM-02": 2,
             "MD-INT-01": 3, "MD-ANC-01": 2, "WT-POP-05": 12, "WT-POP-07": 3, "WT-DFX-01": 4, "WT-FZA-01": 4,
             "WT-FZA-02": 2, "WT-WTP-01": 3, "WT-WTP-02": 2, "WT-WTP-03": 1, "WT-WTP-04": 3}
    for rid, base in picks.items():
        res = C.RESOURCE_BY_ID[rid]
        coll = res["collection"]
        for ln in C.lines(res):
            rk = C.row_key(rid, ln["line"])
            if res["row_type"] == "RANK_SELECT":
                a[(coll, rk, "COLC")] = ["Inspector", "Sergeant", "Constable"][min(ln["line"], 3) - 1]
            for risk, rs in RISK_SCALE.items():
                for d in _days(res):
                    n = max(0, round(base * rs * DAY_SCALE.get(d, 1) / ln["line"]))
                    v[(coll, rk, f"{risk}|{d}|NUM")] = n
                    if n == 0:
                        continue
                    if C.applicable(res, "TOD", d if coll == "MD" else None):
                        v[(coll, rk, f"{risk}|{d}|TOD")] = 10
                    v[(coll, rk, f"{risk}|{d}|XH")] = 3
    # role profile and additional rows
    for coll, rid, rank, role in (("MD", "MD-AAP-01", "Constable", "ARV"), ("MD", "MD-AAP-02", "Sergeant", "CTSFO"),
                                  ("WT", "WT-APO-03", "Constable", "Medic")):
        rk = f"{rid}|1"
        a[(coll, rk, "COLA")], a[(coll, rk, "COLC")] = rank, role
        for risk, rs in RISK_SCALE.items():
            d = "MD" if coll == "MD" else "WT"
            v[(coll, rk, f"{risk}|{d}|NUM")] = round(4 * rs)
            if C.RESOURCE_BY_ID[rid]["supply"] == "FS":
                v[(coll, rk, f"{risk}|{d}|TOD")] = 10
            v[(coll, rk, f"{risk}|{d}|XH")] = 2
    for coll, rid in (("MD", "MD-ADD-01"), ("WT", "WT-ADD-02")):
        rk = f"{rid}|1"
        a[(coll, rk, "COLB")] = "Hotel security liaison (synthetic)"
        a[(coll, rk, "COLC")] = "Sergeant"
        for risk in RISK_SCALE:
            d = "MD" if coll == "MD" else "WT"
            v[(coll, rk, f"{risk}|{d}|NUM")] = 2
            v[(coll, rk, f"{risk}|{d}|XH")] = 3
    if with_errors:
        status["MD-CBR"] = "In progress"                 # DQ12
        v[("MD", "MD-CBR-03|1", "H|MD|NUM")] = 2         # DQ09 rank + DQ10 role missing
        v[("MD", "MD-CBR-03|1", "H|MD|XH")] = 4
        v[("MD", "MD-CUS-01|2", "SH|MD|XH")] = 26        # DQ05 hours > 24
        v[("MD", "MD-INT-02|1", "L|MD|TOD")] = 8         # DQ06 value in N/A cell (MA ToD)
        v[("MD", "MD-ANC-02|1", "M|MD|NUM")] = 3         # DQ07 number without hours
    cover = {"cv_Force": FORCE, "cv_ContactName": "Synthetic Tester", "cv_ContactEmail": "synthetic@example.invalid",
             "cv_SubmissionNo": 1, "cv_SubmissionDate": dt.datetime(2026, 11, 4),
             "cv_Comments": "SYNTHETIC DATA FOR TESTING ONLY - NOT A FORECAST"}
    return cover, status, v, a


def fill_fixtures(path, fixtures):
    """Enter fixture risk levels on Demand Sheet row 4 of a backup calculation workbook."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb["Demand Sheet"]
    names = dict((r, n) for r, n, _ in C.RISK_LEVELS)
    for d, risk in fixtures:
        k = (d - build_calc_workbook.CAL_START).days
        ws.cell(4, build_calc_workbook.FIRST_DATE_COL + k * build_calc_workbook.BLOCK, names[risk])
    wb.save(path)


def build(out_dir, with_errors=True):
    os.makedirs(out_dir, exist_ok=True)
    cover, status, v, a = sample_inputs(with_errors)
    tpl = os.path.join(out_dir, "_tpl.xlsx")
    build_template.build(tpl)
    sub = fill(tpl, os.path.join(out_dir, "SYNTHETIC_Submission_TEST_FORCE.xlsx"),
               cover=cover, status=status, values=v, attrs=a)
    calc_tpl = os.path.join(out_dir, "_calc.xlsx")
    build_calc_workbook.build(calc_tpl)
    calc = fill(calc_tpl, os.path.join(out_dir, "SYNTHETIC_Backup_Calculation_TEST_FORCE.xlsx"),
                cover=cover, status=status, values=v, attrs=a)
    fill_fixtures(calc, SYNTHETIC_FIXTURES)
    for f in (tpl, calc_tpl):
        os.remove(f)
    return sub, calc


if __name__ == "__main__":
    for p in build(os.path.abspath(os.path.join(HERE, "..", "output", "samples"))):
        print(p)
