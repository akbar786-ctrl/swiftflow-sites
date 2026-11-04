"""
Builds a SYNTHETIC completed submission for Power BI development and testing.
All figures are invented test values - they are NOT forecasts and must not be
reported. The force name is deliberately not a real force. It also contains
deliberate data-quality problems so the DQ report can be exercised.
"""
import datetime as dt
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "tests"))
import build_template  # noqa: E402
import catalogue as C  # noqa: E402
from fill import fill  # noqa: E402

RISK_SCALE = {"L": 1.0, "M": 1.3, "H": 1.7, "SH": 2.2}
DAY_SCALE = {"MD-2": 0.3, "MD-1": 0.6, "MD": 1.0, "MD+1": 0.4}


def build(template, out):
    v, a = {}, {}
    status = {f"{c}-{s}": "Complete" for (c, s) in C.SECTIONS}
    status["MD-CBR"] = "In progress"                      # DQ12 example
    # PSU configurations: 1 Insp, 3 Sgt, 18 PC per L2 PSU; 1 Sgt, 7 PC per L3 serial
    for res in C.RESOURCES:
        if res["row_type"] != "PSU_CONFIG":
            continue
        comp = {"INSP": 1, "SGT": 3, "PC": 18} if "Level 2" in res["name"] else {"INSP": 0, "SGT": 1, "PC": 7}
        days = ["MD"] if res["matchday_only"] else (["WT"] if res["collection"] == "WT" else list(DAY_SCALE))
        for ln in C.lines(res):
            for risk in RISK_SCALE:
                for d in days:
                    rk = C.row_key(res["resource_id"], ln["line"])
                    v[(res["collection"], rk, f"{risk}|{d}|NUM")] = comp[ln["rank"]]
                    v[(res["collection"], rk, f"{risk}|{d}|TOD")] = 10
                    v[(res["collection"], rk, f"{risk}|{d}|XH")] = 2 if d != "MD" else 4
    # PSU counts (force + mutual aid)
    for rid, base in [("MD-POP-01", 2), ("MD-POP-02", 1), ("MD-POP-03", 4), ("MD-POP-04", 2), ("MD-POP-05", 6),
                      ("MD-POP-06", 3), ("MD-POP-07", 2), ("MD-POP-08", 1)]:
        res = C.RESOURCE_BY_ID[rid]
        days = ["MD"] if res["matchday_only"] else list(DAY_SCALE)
        for ln in C.lines(res):
            for risk, rs in RISK_SCALE.items():
                for d in days:
                    n = round(base * rs * (DAY_SCALE[d] if d != "MD" else 1) / ln["line"])
                    v[("MD", C.row_key(rid, ln["line"]), f"{risk}|{d}|NUM")] = n
    for rid, base in [("WT-POP-01", 2), ("WT-POP-02", 1), ("WT-POP-03", 3), ("WT-POP-04", 1)]:
        for ln in C.lines(C.RESOURCE_BY_ID[rid]):
            for risk, rs in RISK_SCALE.items():
                v[("WT", C.row_key(rid, ln["line"]), f"{risk}|WT|NUM")] = round(base * rs / ln["line"])
    # A selection of officer resources
    picks = {"MD-EGT-01": 4, "MD-EGT-02": 2, "MD-MNT-03": 8, "MD-DOG-05": 4, "MD-FBP-01": 6, "MD-FBP-02": 4,
             "MD-RDS-01": 10, "MD-ARM-03": 6, "MD-ARM-05": 4, "MD-SRC-01": 3, "MD-UAS-03": 2, "MD-LFR-03": 2,
             "MD-CBR-01": 3, "MD-CMD-01": 1, "MD-CTL-01": 5, "MD-CUS-01": 4, "MD-CRM-01": 3, "MD-INT-01": 3,
             "MD-ANC-01": 2, "WT-POP-05": 12, "WT-POP-07": 3, "WT-DFX-01": 4, "WT-FZA-01": 4, "WT-WTP-01": 3,
             "WT-WTP-02": 2, "WT-WTP-04": 3}
    for rid, base in picks.items():
        res = C.RESOURCE_BY_ID[rid]
        days = list(DAY_SCALE) if res["collection"] == "MD" else ["WT"]
        for ln in C.lines(res):
            rk = C.row_key(rid, ln["line"])
            if res["row_type"] == "RANK_SELECT":
                a[(res["collection"], rk, "RANK")] = ["Inspector", "Sergeant", "Constable"][min(ln["line"], 3) - 1]
            for risk, rs in RISK_SCALE.items():
                for d in days:
                    n = max(0, round(base * rs * DAY_SCALE.get(d, 1) / ln["line"]))
                    v[(res["collection"], rk, f"{risk}|{d}|NUM")] = n
                    if n == 0:
                        continue
                    if C.applicable(res, "TOD", d):
                        v[(res["collection"], rk, f"{risk}|{d}|TOD")] = 10
                    v[(res["collection"], rk, f"{risk}|{d}|XH")] = 3
    # Role profile and additional rows
    a[("MD", "MD-AAP-01|1", "ROLE")] = "ARV"
    a[("MD", "MD-AAP-01|1", "RANK")] = "Constable"
    a[("MD", "MD-ADD-01|1", "ROLE")] = "Hotel security liaison (synthetic)"
    a[("MD", "MD-ADD-01|1", "SUPPLY")] = "Force Supplied"
    a[("MD", "MD-ADD-01|1", "RANK")] = "Sergeant"
    for risk, rs in RISK_SCALE.items():
        v[("MD", "MD-AAP-01|1", f"{risk}|MD|NUM")] = round(6 * rs)
        v[("MD", "MD-AAP-01|1", f"{risk}|MD|TOD")] = 10
        v[("MD", "MD-AAP-01|1", f"{risk}|MD|XH")] = 2
        v[("MD", "MD-ADD-01|1", f"{risk}|MD|NUM")] = 2
        v[("MD", "MD-ADD-01|1", f"{risk}|MD|TOD")] = 8
    # Deliberate data-quality problems
    v[("MD", "MD-CBR-03|1", "H|MD|NUM")] = 2          # DQ09 rank + DQ10 role missing
    v[("MD", "MD-CBR-03|1", "H|MD|XH")] = 4
    v[("MD", "MD-CUS-01|2", "SH|MD|XH")] = 26         # DQ05 hours > 24
    v[("MD", "MD-INT-02|1", "L|MD|TOD")] = 8          # DQ06 value in N/A cell (MA ToD)
    v[("MD", "MD-ANC-02|1", "M|MD|NUM")] = 3          # DQ07 number without hours
    cover = {"cv_Force": "SYNTHETIC TEST FORCE", "cv_ContactName": "Synthetic Tester",
             "cv_ContactEmail": "synthetic@example.invalid", "cv_SubmissionNo": 1,
             "cv_SubmissionDate": dt.datetime(2026, 11, 4),
             "cv_Comments": "SYNTHETIC DATA FOR POWER BI TESTING ONLY - NOT A FORECAST"}
    return fill(template, out, cover=cover, status=status, values=v, attrs=a)


if __name__ == "__main__":
    out_dir = os.path.abspath(os.path.join(HERE, "..", "output", "samples"))
    os.makedirs(out_dir, exist_ok=True)
    tpl = os.path.join(HERE, "..", "output", f"E28_Collection_Template_v{C.TEMPLATE_VERSION}.xlsx")
    if not os.path.exists(tpl):
        build_template.build(tpl)
    print(build(tpl, os.path.join(out_dir, "SYNTHETIC_Submission_TEST_FORCE.xlsx")))
