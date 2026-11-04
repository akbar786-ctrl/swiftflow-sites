"""
EURO 2028 reference calculation engine.

Plain-Python implementation of exactly the rules implemented by
  * the Power BI model (powerbi/*.pq, powerbi/measures.dax), and
  * the backup calculation workbook (build/build_calc_workbook.py).
It exists so the rules can be unit-tested and the two implementations
reconciled against an independent one (docs/04_validation.md).

Two rule sets (parameter RuleSet):
  LEGACY  reproduces the old workbook's logic (default; selected by the project team)
  AGREED  the revised rules recorded as decisions D-02, D-03, D-05, D-06, D-07
"""
import datetime as dt
import os
import sys
from collections import defaultdict

import openpyxl

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_HERE, "..", "build"))
import catalogue as C  # noqa: E402
import reference_data as R  # noqa: E402

SEV_ERROR, SEV_WARN, SEV_INFO = "Error", "Warning", "Info"
OFFICER_TYPES = ("OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL")
DAY_INDEX = {"MD-2": 1, "MD-1": 2, "MD": 3, "MD+1": 4}
OFFSET = {d: o for d, _, o, _ in C.MD_DAYS}
# Legacy priority of relative days when a date is several relative days for
# fixtures of the same risk level (verified against every legacy Demand Sheet row).
PRIORITY = {"STD": ["MD", "MD-1", "MD-2", "MD+1"], "NONMD": ["MD-1", "MD-2", "MD+1"], "MDONLY": ["MD"]}


# ---------------------------------------------------------------------------
# Rates
# ---------------------------------------------------------------------------
class Rates:
    def __init__(self, params):
        self.p = params
        self.effective = params["RateEffectiveDate"]
        self.base = {(t, rk): (tod, ot) for t, rk, tod, ot, *_ in R.BASE_RATES}
        self.zero = {(s, rk) for s, rk, *_ in R.ZERO_COST_RULES}
        self.package_base = {u: v for u, v, _ in R.MA_PSU_PACKAGE}
        self.legacy = params["RuleSet"] == "LEGACY"

    def factor(self, table):
        f = 1.0
        for t, eff, fac, _ in R.UPLIFT_STEPS:
            if t == table and eff <= self.effective:
                f *= fac
        return f

    def table_rate(self, table, rate_rank):
        b = self.base.get((table, rate_rank))
        return None if b is None else (b[0] * self.factor(table), b[1] * self.factor(table))

    def hourly(self, supply, rank, res=None):
        """(tod_rate, ot_rate), or None when no rate exists."""
        if res is not None and res["row_type"] == "ADDITIONAL" and self.legacy:
            # Legacy Cost Calculator: MD additional rows at MA Sergeant, WT additional rows at HF Constable
            t, rk = ("MA", "SGT") if res["collection"] == "MD" else ("HF", "PC")
            return 0.0, self.table_rate(t, rk)[1]
        if rank is None or supply is None:
            return None
        rate_rank = next((r[2] for r in C.RANKS if r[0] == rank), rank)
        if self.legacy and res is not None and res.get("legacy_insp_ma") and supply == "FS" and rate_rank == "INSP":
            return 0.0, self.table_rate("MA", "INSP")[1]   # legacy: In-Stadium PSU Inspectors at MA rate
        if (supply, rate_rank) in self.zero:
            return 0.0, 0.0
        r = self.table_rate("HF" if supply == "FS" else "MA", rate_rank)
        if r is None:
            return None
        return (0.0 if self.legacy else r[0]), r[1]       # legacy: Extra Hours only

    def package(self, unit_label):
        b = self.package_base.get(unit_label)
        return None if b is None else b * self.factor("MAPSU")


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------
def _num(v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return None, "BLANK"
    if isinstance(v, bool):
        return None, "INVALID"
    if isinstance(v, (int, float)):
        return float(v), "VALUE"
    try:
        return float(str(v).strip()), "VALUE"
    except ValueError:
        return None, "INVALID"


def _txt(v):
    return str(v).strip() if v not in (None, "") else None


class Submission:
    def __init__(self, path):
        self.path, self.file = path, os.path.basename(path)
        self.cover, self.section_status, self.cells, self.lines = {}, {}, [], {}
        self.structure_issues, self.optional = [], {}
        self.submission_id = None


def default_params(**overrides):
    p = {x[0]: x[1] for x in R.PARAMETERS}
    p.update(overrides)
    return p


class Engine:
    def __init__(self, params=None):
        self.params = default_params(**(params or {}))
        self.legacy = self.params["RuleSet"] == "LEGACY"
        self.rates = Rates(self.params)
        self.ma_tod = str(self.params["MA_ToD_Collected"]).upper() == "TRUE"
        self.expected_keys = {coll: {C.row_key(r["resource_id"], ln["line"]) for r in C.RESOURCES
                                     if r["collection"] == coll for ln in C.lines(r)} for coll in ("MD", "WT")}

    def read_submission(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        sub = Submission(path)
        for key in ("cv_Force", "cv_ContactName", "cv_ContactEmail", "cv_SubmissionDate", "cv_SubmissionNo",
                    "cv_TemplateVersion", "cv_Comments"):
            if key in wb.defined_names:
                for sheet, ref in wb.defined_names[key].destinations:
                    sub.cover[key[3:]] = wb[sheet][ref.replace("$", "")].value
        sub.submission_id = f"{sub.cover.get('Force') or 'UNKNOWN'}#{sub.cover.get('SubmissionNo') or 0}"
        if "Cover" in wb.sheetnames and "tblSectionStatus" in wb["Cover"].tables:
            rows = list(wb["Cover"][wb["Cover"].tables["tblSectionStatus"].ref])
            hdr = [c.value for c in rows[0]]
            for rw in rows[1:]:
                d = dict(zip(hdr, [c.value for c in rw]))
                sub.section_status[d["SectionKey"]] = d["Status"] or "Not started"
        for coll, sheet in C.SHEET_NAMES.items():
            if sheet not in wb.sheetnames:
                sub.structure_issues.append((coll, None, None, "DQ03", SEV_ERROR, f"Sheet '{sheet}' missing"))
                continue
            self._read_sheet(wb[sheet], coll, sub)
        for sheet, tname in (("Command Team", "tblCommandTeam"), ("Other Resources", "tblOtherResources")):
            if sheet in wb.sheetnames and tname in wb[sheet].tables:
                rows = list(wb[sheet][wb[sheet].tables[tname].ref])
                hdr = [c.value for c in rows[0]]
                sub.optional[tname] = [dict(zip(hdr, [c.value for c in rw])) for rw in rows[1:]
                                       if any(c.value not in (None, "") for c in rw[2:4])]
        return sub

    def _read_sheet(self, ws, coll, sub):
        key_row = keycol = None
        for r in range(1, 15):
            for c in range(1, ws.max_column + 1):
                if ws.cell(r, c).value == "ROWKEY":
                    key_row, keycol = r, c
                    break
            if key_row:
                break
        if key_row is None:
            sub.structure_issues.append((coll, None, None, "DQ03", SEV_ERROR, "Key row not found"))
            return
        keys = {ws.cell(key_row, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(key_row, c).value}
        inputs = {c: k.split("|") for k, c in keys.items() if k.count("|") == 2 and not k.startswith("TOTAL")}
        seen = set()
        for r in range(key_row + 1, ws.max_row + 1):
            rk = ws.cell(r, keycol).value
            if not rk or rk.startswith("SECTION|"):
                continue
            seen.add(rk)
            if rk not in self.expected_keys[coll]:
                sub.structure_issues.append((coll, rk, None, "DQ03", SEV_ERROR, "Unknown row key (rows changed?)"))
                continue
            rid, line = rk.rsplit("|", 1)
            res = C.RESOURCE_BY_ID[rid]
            ln = next(x for x in C.lines(res) if x["line"] == int(line))
            a, b, c = (_txt(ws.cell(r, keys[k]).value) for k in ("COLA", "COLB", "COLC"))
            rt = res["row_type"]
            at = dict(resource_id=rid, line=int(line), res=res, supply=res["supply"], location=res["location"],
                      unit_label=None, role=None, description=None, rank=None)
            if rt in ("PSU_COUNT", "MA_PSU"):
                at["unit_label"] = c or ln["label"]
            elif rt == "ROLE":
                at["rank"], at["role"] = C.RANK_BY_LABEL.get(a), c
            elif ln["rank"]:
                at["rank"] = ln["rank"]
            else:
                at["rank"] = C.RANK_BY_LABEL.get(c)
            if rt == "ADDITIONAL":
                at["description"] = b
            sub.lines[rk] = at
            for col, (risk, day, meas) in inputs.items():
                val, state = _num(ws.cell(r, col).value)
                sub.cells.append(dict(row_key=rk, risk=risk, day=day, measure=meas, value=val, state=state,
                                      raw=ws.cell(r, col).value, cell=ws.cell(r, col).coordinate))
        for missing in sorted(self.expected_keys[coll] - seen):
            sub.structure_issues.append((coll, missing, None, "DQ03", SEV_ERROR, "Expected row missing"))

    # ---- calculation -------------------------------------------------------
    def calculate(self, submissions, fixtures=None):
        subs = self._latest_per_force(submissions)
        dq, inputs = [], []
        for s in submissions:
            if s not in subs:
                dq.append(self._dq(s, None, None, "DQ16", SEV_INFO, "Superseded by a later submission"))
        for s in subs:
            dq.extend(self._dq(s, i[0], i[1], i[3], i[4], i[5], i[2]) for i in s.structure_issues)
            dq.extend(self._check_cover(s))
            inputs.extend(self._clean_inputs(s, dq))
        deployed, packages = self._deploy(inputs, dq)
        self._cost(deployed, dq)
        return dict(inputs=inputs, deployed=deployed, packages=packages, dq=dq,
                    wt_tournament=self._wt_tournament(deployed, packages),
                    calendar=self.expand_calendar(deployed, packages, fixtures or []))

    @staticmethod
    def _latest_per_force(subs):
        best = {}
        for s in subs:
            f, n = s.cover.get("Force") or s.file, s.cover.get("SubmissionNo") or 0
            if f not in best or n > (best[f].cover.get("SubmissionNo") or 0):
                best[f] = s
        return list(best.values())

    @staticmethod
    def _dq(sub, coll, row_key, code, sev, detail, cell=None):
        return dict(submission=sub.submission_id, force=sub.cover.get("Force"), file=sub.file, collection=coll,
                    row_key=row_key, cell=cell, code=code, severity=sev, detail=detail)

    def _check_cover(self, s):
        out = []
        if not s.cover.get("Force"):
            out.append(self._dq(s, None, None, "DQ01", SEV_ERROR, "Force not selected on Cover"))
        if str(s.cover.get("TemplateVersion")) != self.params["TemplateVersion"]:
            out.append(self._dq(s, None, None, "DQ02", SEV_ERROR,
                                f"Template version {s.cover.get('TemplateVersion')} not accepted"))
        for key in ("ContactName", "ContactEmail"):
            if not s.cover.get(key):
                out.append(self._dq(s, None, None, "DQ01", SEV_WARN, f"Cover field {key} missing"))
        return out

    def _clean_inputs(self, s, dq):
        grouped = defaultdict(dict)
        for c in s.cells:
            ln = s.lines[c["row_key"]]
            res = ln["res"]
            sec = f"{res['collection']}-{res['section_code']}"
            status = s.section_status.get(sec, "Not started")
            applicable = C.applicable(res, c["measure"], c["day"] if res["collection"] == "MD" else None)
            if self.ma_tod and res["supply"] == "MA" and c["measure"] == "TOD":
                # OI-01 switch: collect MA Normal ToD wherever Extra Hours are collected
                applicable = C.applicable(res, "XH", c["day"] if res["collection"] == "MD" else None)
            val, state = c["value"], c["state"]
            if state == "INVALID":
                dq.append(self._dq(s, res["collection"], c["row_key"], "DQ04", SEV_ERROR,
                                   f"Non-numeric value '{c['raw']}' - excluded", c["cell"]))
            if not applicable:
                if state == "VALUE":
                    dq.append(self._dq(s, res["collection"], c["row_key"], "DQ06", SEV_WARN,
                                       "Value in a not-applicable cell - excluded", c["cell"]))
                val, state = None, "NA"
            elif state == "VALUE":
                bad = None
                if val < 0:
                    bad = ("DQ04", "Negative value")
                elif c["measure"] == "NUM" and val != int(val):
                    bad = ("DQ04", "Number required must be a whole number")
                elif c["measure"] in ("TOD", "XH") and val > 24:
                    bad = ("DQ05", "Hours above 24")
                if bad:
                    dq.append(self._dq(s, res["collection"], c["row_key"], bad[0], SEV_ERROR,
                                       f"{bad[1]} - excluded", c["cell"]))
                    val, state = None, "INVALID"
            elif state == "BLANK" and status in ("Complete", "Not required"):
                val, state = 0.0, "BLANK_AS_ZERO"
            elif state == "INVALID":
                val = None
            grouped[(c["row_key"], c["risk"], c["day"])][c["measure"]] = (val, state)
        out = []
        for (rk, risk, day), m in grouped.items():
            ln = s.lines[rk]
            res = ln["res"]
            rec = dict(submission=s.submission_id, force=s.cover.get("Force"), collection=res["collection"],
                       row_key=rk, resource_id=res["resource_id"], line=ln["line"], row_type=res["row_type"],
                       section=f"{res['collection']}-{res['section_code']}", portfolio=res["portfolio"],
                       resource=res["name"], supply=ln["supply"], rank=ln["rank"], role=ln["role"],
                       description=ln["description"], location=ln["location"], unit_label=ln["unit_label"],
                       day_group=res["day_group"], risk=risk, day=day,
                       num=m.get("NUM", (None, "NA"))[0], tod=m.get("TOD", (None, "NA"))[0],
                       xh=m.get("XH", (None, "NA"))[0], num_state=m.get("NUM", (None, "NA"))[1],
                       status=s.section_status.get(f"{res['collection']}-{res['section_code']}", "Not started"))
            self._row_checks(s, rec, dq)
            out.append(rec)
        for sec, st in s.section_status.items():
            if st not in ("Complete", "Not required"):
                dq.append(self._dq(s, sec[:2], f"SECTION|{sec}", "DQ12", SEV_WARN,
                                   f"Section '{C.SECTIONS.get((sec[:2], sec[3:]), sec)}' status: {st}"))
        return out

    def _row_checks(self, s, rec, dq):
        rt, ctx = rec["row_type"], f"{rec['risk']} {rec['day']}: "
        has = (rec["num"] or 0) > 0
        add = lambda code, sev, msg: dq.append(self._dq(s, rec["collection"], rec["row_key"], code, sev, ctx + msg))  # noqa: E731
        if rt in OFFICER_TYPES and has:
            if not rec["tod"] and not rec["xh"]:
                add("DQ07", SEV_WARN, "number entered but no hours")
            if (rec["tod"] or 0) + (rec["xh"] or 0) > 24:
                add("DQ05", SEV_ERROR, "Normal ToD + Extra Hours above 24")
            if not rec["rank"]:
                add("DQ09", SEV_ERROR if not (self.legacy and rt == "ADDITIONAL") else SEV_WARN,
                    "rank not selected" + ("" if (self.legacy and rt == "ADDITIONAL") else " - not costed"))
            if rt == "ROLE" and not rec["role"]:
                add("DQ10", SEV_WARN, "role profile not selected")
            if rt == "ADDITIONAL" and not rec["description"]:
                add("DQ11", SEV_WARN, "additional resource not described")
        if rt in OFFICER_TYPES and not has and ((rec["tod"] or 0) > 0 or (rec["xh"] or 0) > 0):
            add("DQ08", SEV_WARN, "hours entered but no number")

    # ---- deployment ----------------------------------------------------------
    def _deploy(self, inputs, dq):
        by_key = {(i["submission"], i["resource_id"], i["line"], i["risk"], i["day"]): i for i in inputs}
        deployed, packages = [], []
        md1 = 1.0 if self.legacy else float(self.params["MD1_PSU_Factor"])
        for i in inputs:
            rt = i["row_type"]
            base = {k: i[k] for k in ("submission", "force", "collection", "resource_id", "resource", "portfolio",
                                      "section", "supply", "role", "location", "risk", "day", "day_group", "row_key")}
            if rt in OFFICER_TYPES:
                n = i["num"] or 0
                if n > 0:
                    deployed.append(dict(base, rank=i["rank"], source="Direct", units=None, unit_label=None,
                                         line_key=i["row_key"], officers=n, tod_hours=n * (i["tod"] or 0),
                                         xh_hours=n * (i["xh"] or 0)))
            elif rt == "PSU_COUNT":
                units = (i["num"] or 0) * (md1 if i["day"] == "MD+1" else 1)
                if units <= 0:
                    continue
                res = C.RESOURCE_BY_ID[i["resource_id"]]
                cfg = C.RESOURCE_BY_ID[res["psu_config"][res["units"][i["line"] - 1]]]
                found = False
                for cl in C.lines(cfg):
                    c = by_key.get((i["submission"], cfg["resource_id"], cl["line"], i["risk"], i["day"])) or {}
                    per = c.get("num") or 0
                    if per <= 0:
                        continue
                    found = True
                    off = units * per
                    deployed.append(dict(base, rank=cl["rank"], source="PSU expansion", units=units,
                                         unit_label=i["unit_label"], config_id=cfg["resource_id"],
                                         line_key=f"{i['row_key']}|{cl['rank']}", officers=off,
                                         tod_hours=off * (c.get("tod") or 0), xh_hours=off * (c.get("xh") or 0)))
                if not found:
                    dq.append(dict(submission=i["submission"], force=i["force"], file=None,
                                   collection=i["collection"], row_key=i["row_key"], cell=None, code="DQ13",
                                   severity=SEV_ERROR, detail=f"{i['risk']} {i['day']}: {units:g} units entered "
                                                              f"but PSU configuration '{cfg['name']}' is empty"))
        # mutual-aid PSU packages (day grain; PackagesCharged allocated per basis)
        grp = defaultdict(list)
        for i in inputs:
            if i["row_type"] == "MA_PSU" and (i["num"] or 0) > 0:
                grp[(i["submission"], i["resource_id"], i["line"], i["risk"])].append(i)
        for (sub, rid, line, risk), rows in grp.items():
            res = C.RESOURCE_BY_ID[rid]
            if self.legacy or res["collection"] == "WT":
                basis = "PER_UNIT_PER_DAY" if self.legacy else self.params["MAPSU_Basis_WT"]
            elif res["matchday_only"]:
                basis = self.params["MAPSU_Basis_Matchday"]
            else:
                basis = self.params["MAPSU_Basis_NonMatchday"]
            peak_row = max(rows, key=lambda r: (r["num"], -DAY_INDEX.get(r["day"], 0)))
            unit = rows[0]["unit_label"]
            rate = self.rates.package(unit)
            if rate is None:
                dq.append(dict(submission=sub, force=rows[0]["force"], file=None, collection=rows[0]["collection"],
                               row_key=rows[0]["row_key"], cell=None, code="DQ14", severity=SEV_WARN,
                               detail=f"No mutual-aid package rate for '{unit}' - not costed (OI-03)"))
            for r in rows:
                if basis.endswith("MAXDAY"):
                    charged = r["num"] if r is peak_row else 0
                elif basis == "PER_UNIT_PER_FIXTURE":
                    charged = r["num"] if r["day"] == "MD" else 0
                else:
                    charged = r["num"]
                packages.append(dict(submission=sub, force=r["force"], collection=r["collection"], resource_id=rid,
                                     resource=res["name"], portfolio=res["portfolio"], location=res["location"],
                                     day_group=res["day_group"], row_key=r["row_key"], line_key=r["row_key"],
                                     risk=risk, day=r["day"], unit_label=unit, basis=basis, units=r["num"],
                                     packages=charged, package_rate=rate,
                                     cost=(charged * rate) if rate is not None else None))
        return deployed, packages

    def _cost(self, deployed, dq):
        for d in deployed:
            res = C.RESOURCE_BY_ID[d["config_id"]] if d.get("config_id") else C.RESOURCE_BY_ID[d["resource_id"]]
            rates = self.rates.hourly(d["supply"], d["rank"], res)
            if rates is None:
                d.update(tod_rate=None, ot_rate=None, tod_cost=None, xh_cost=None, cost=None)
                if d["rank"]:
                    dq.append(dict(submission=d["submission"], force=d["force"], file=None,
                                   collection=d["collection"], row_key=d["row_key"], cell=None, code="DQ14",
                                   severity=SEV_WARN, detail=f"No {d['supply']} rate for rank {d['rank']} - not costed"))
                continue
            d["tod_rate"], d["ot_rate"] = rates
            d["tod_cost"], d["xh_cost"] = d["tod_hours"] * rates[0], d["xh_hours"] * rates[1]
            d["cost"] = d["tod_cost"] + d["xh_cost"]

    def _wt_tournament(self, deployed, packages):
        days = {r: n for r, n, _ in R.WT_DAY_ASSUMPTIONS}
        out = defaultdict(float)
        for d in list(deployed) + list(packages):
            if d["collection"] == "WT" and d.get("cost") is not None:
                out[(d["force"], d["risk"])] += d["cost"] * days[d["risk"]]
        return dict(out)

    # ---- calendar ------------------------------------------------------------
    def expand_calendar(self, deployed, packages, fixtures):
        """
        fixtures: list of dict(fixture_id, date, host_force, risk).
        AGREED: every fixture's MD-2..MD+1 demand is placed on its dates and coinciding dates are ADDED.
        LEGACY: for each date, risk level and line, only ONE relative day is used - the first found in the
                line's priority group (STD: MD > MD-1 > MD-2 > MD+1; NONMD: MD-1 > MD-2 > MD+1; MDONLY: MD).
                Fixtures at different risk levels are added together, as the legacy templates summed L+M+H.
        """
        rows = [dict(d, kind="officer") for d in deployed if d["collection"] == "MD"] + \
               [dict(p, kind="package") for p in packages if p["collection"] == "MD"]
        cal = []
        if not self.legacy:
            for fx in fixtures:
                for d in rows:
                    if d["force"] == fx["host_force"] and d["risk"] == fx["risk"]:
                        cal.append(dict(d, fixture_id=fx["fixture_id"], date=fx["date"] + dt.timedelta(OFFSET[d["day"]])))
            return cal
        by_line = defaultdict(dict)        # (force, risk, kind, line_key) -> {day: row}
        for d in rows:
            by_line[(d["force"], d["risk"], d["kind"], d["line_key"])][d["day"]] = d
        fx_by = defaultdict(dict)          # (force, risk) -> {date: fixture_id}
        for fx in fixtures:
            fx_by[(fx["host_force"], fx["risk"])][fx["date"]] = fx["fixture_id"]
        for (force, risk, kind, lk), days in by_line.items():
            fdates = fx_by.get((force, risk), {})
            if not fdates:
                continue
            group = next(iter(days.values()))["day_group"]
            cand = set()
            for f in fdates:
                for off in (-2, -1, 0, 1):
                    cand.add(f + dt.timedelta(off))
            for t in sorted(cand):
                for day in PRIORITY[group]:
                    fdate = t - dt.timedelta(OFFSET[day])
                    if fdate in fdates:
                        if day in days:
                            cal.append(dict(days[day], fixture_id=fdates[fdate], date=t))
                        break
        return cal


# ---------------------------------------------------------------------------
# Command Team & Planning, Other Resources (legacy methods)
# ---------------------------------------------------------------------------
def months_between(first=dt.date(2026, 9, 1), last=dt.date(2029, 3, 1)):
    out, d = [], first
    while d <= last:
        out.append(d)
        d = dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return out


def command_team_costs(posts, params=None):
    """
    posts: iterable of dict(rank, posts, start, end) - rank code, number of posts, first and LAST month (inclusive).
    Monthly cost per post = annual CTP salary / 12 x uplift factor at that month (legacy method).
    Returns list of dict(rank, month, cost).
    """
    p = default_params(**(params or {}))
    ann = {rk: v for t, rk, v in R.ANNUAL_COSTS if t == "CTP"}
    out = []
    for row in posts:
        n, start, end = row.get("posts") or 0, row.get("start"), row.get("end")
        if not n or not start or row["rank"] not in ann:
            continue
        for m in months_between():
            if m >= start and (end is None or m <= end):
                f = 1.0
                for t, eff, fac, _ in R.UPLIFT_STEPS:
                    if t == "CTP" and eff <= m:
                        f *= fac
                out.append(dict(rank=row["rank"], month=m, cost=n * ann[row["rank"]] / 12 * f))
    return out


def other_resources_costs(rows, params=None):
    """rows: iterable of dict(category, rank, hours_2728, hours_2829). Priced at HF overtime rates."""
    p = default_params(**(params or {}))
    r27, r28 = Rates(dict(p, RateEffectiveDate=p["OtherResourcesRateDate2728"])), Rates(p)
    out = []
    for row in rows:
        rk = row["rank"]
        a27, a28 = r27.table_rate("HF", rk), r28.table_rate("HF", rk)
        if a27 is None:
            continue
        out.append(dict(row, cost_2728=(row.get("hours_2728") or 0) * a27[1],
                        cost_2829=(row.get("hours_2829") or 0) * a28[1]))
    return out


def summarise(rows, keys, value="cost"):
    out = defaultdict(float)
    for r in rows:
        if r.get(value) is not None:
            out[tuple(r[k] for k in keys)] += r[value]
    return dict(out)
