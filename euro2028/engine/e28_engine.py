"""
EURO 2028 reference calculation engine.

A plain-Python implementation of exactly the rules the Power BI model
implements (powerbi/*.pq and powerbi/measures.dax). It exists so that the
rules can be unit-tested, and so that Power BI outputs can be reconciled
against an independent implementation (see docs/04_validation.md).

Usage:
    from e28_engine import Engine
    eng = Engine()                      # loads catalogue + reference data
    sub = eng.read_submission("force.xlsx")
    res = eng.calculate([sub])          # dict of fact tables + DQ issues
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


# ---------------------------------------------------------------------------
# Rates
# ---------------------------------------------------------------------------
class Rates:
    def __init__(self, effective=None, params=None):
        params = params or {p[0]: p[1] for p in R.PARAMETERS}
        self.effective = effective or params["RateEffectiveDate"]
        self.base = {(t, rk): (tod, ot) for t, rk, tod, ot, *_ in R.BASE_RATES}
        self.zero = {(s, rk) for s, rk, *_ in R.ZERO_COST_RULES}
        self.package_base = {u: v for u, v, _ in R.MA_PSU_PACKAGE}

    def factor(self, table):
        f = 1.0
        for t, eff, fac, _ in R.UPLIFT_STEPS:
            if t == table and eff <= self.effective:
                f *= fac
        return f

    def hourly(self, supply, rank):
        """(tod_rate, ot_rate) or None if no rate exists."""
        rate_rank = next((r[2] for r in C.RANKS if r[0] == rank), rank)
        if (supply, rate_rank) in self.zero:
            return 0.0, 0.0
        table = "HF" if supply == "FS" else "MA"
        b = self.base.get((table, rate_rank))
        if b is None:
            return None
        f = self.factor(table)
        return b[0] * f, b[1] * f

    def package(self, unit_label):
        b = self.package_base.get(unit_label)
        return None if b is None else b * self.factor("MAPSU")


# ---------------------------------------------------------------------------
# Submission reader
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


class Submission:
    def __init__(self, path):
        self.path = path
        self.file = os.path.basename(path)
        self.cover = {}
        self.section_status = {}
        self.cells = []          # dicts: one per input cell
        self.lines = {}          # row key -> line attributes
        self.structure_issues = []
        self.optional = {}


class Engine:
    def __init__(self, params=None):
        self.params = {p[0]: p[1] for p in R.PARAMETERS}
        if params:
            self.params.update(params)
        self.rates = Rates(self.params["RateEffectiveDate"], self.params)
        self.ma_tod = str(self.params["MA_ToD_Collected"]).upper() == "TRUE"
        self.expected_keys = {
            coll: {C.row_key(r["resource_id"], ln["line"])
                   for r in C.RESOURCES if r["collection"] == coll for ln in C.lines(r)}
            for coll in ("MD", "WT")}

    # ---- reading -----------------------------------------------------------
    def read_submission(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        sub = Submission(path)
        names = wb.defined_names
        for key in ("cv_Force", "cv_ContactName", "cv_ContactEmail", "cv_SubmissionDate",
                    "cv_SubmissionNo", "cv_TemplateVersion", "cv_Comments"):
            if key in names:
                for sheet, ref in names[key].destinations:
                    sub.cover[key[3:]] = wb[sheet][ref.replace("$", "")].value
        sub.submission_id = f"{sub.cover.get('Force') or 'UNKNOWN'}#{sub.cover.get('SubmissionNo') or 0}"
        for coll in ("MD", "WT"):
            sheet = f"{coll} Collection"
            if sheet not in wb.sheetnames:
                sub.structure_issues.append((coll, None, None, "DQ03", SEV_ERROR,
                                             f"Sheet '{sheet}' missing"))
                continue
            self._read_sheet(wb[sheet], coll, sub)
        for sheet, tname in (("Command Team", "tblCommandTeam"), ("Other Resources", "tblOtherResources")):
            if sheet in wb.sheetnames and tname in wb[sheet].tables:
                ws = wb[sheet]
                ref = ws.tables[tname].ref
                rows = list(ws[ref])
                hdr = [c.value for c in rows[0]]
                sub.optional[tname] = [dict(zip(hdr, [c.value for c in rw])) for rw in rows[1:]
                                       if any(c.value not in (None, "") for c in rw[2:])]
        return sub

    def _read_sheet(self, ws, coll, sub):
        key_row = None
        for r in range(1, 15):
            if ws.cell(r, 1).value == "ROWKEY":
                key_row = r
                break
        if key_row is None:
            sub.structure_issues.append((coll, None, None, "DQ03", SEV_ERROR, "Column key row not found"))
            return
        colkeys = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(key_row, c).value
            if v:
                colkeys[c] = v
        meta = {v: c for c, v in colkeys.items() if "|" not in v}
        inputs = {c: v.split("|") for c, v in colkeys.items() if v.count("|") == 2 and not v.startswith("TOTAL")}
        seen = set()
        for r in range(key_row + 1, ws.max_row + 1):
            rk = ws.cell(r, 1).value
            if not rk:
                continue
            if rk.startswith("SECTION|"):
                sub.section_status[rk[8:]] = ws.cell(r, meta.get("LOCATION", 7)).value or "Not started"
                continue
            seen.add(rk)
            if rk not in self.expected_keys[coll]:
                sub.structure_issues.append((coll, rk, None, "DQ03", SEV_ERROR, "Unknown row key (rows changed?)"))
                continue
            rid, line = rk.rsplit("|", 1)
            res = C.RESOURCE_BY_ID[rid]
            ln = next(x for x in C.lines(res) if x["line"] == int(line))
            rank_txt = ws.cell(r, meta["RANK"]).value
            supply_txt = ws.cell(r, meta["SUPPLY"]).value
            attrs = dict(
                resource_id=rid, line=int(line), res=res, collection=coll,
                rank_text=(str(rank_txt).strip() if rank_txt else None),
                role=(str(ws.cell(r, meta["ROLE"]).value).strip() if ws.cell(r, meta["ROLE"]).value else None),
                supply_text=supply_txt,
                location=ws.cell(r, meta["LOCATION"]).value,
            )
            if res["row_type"] in ("PSU_COUNT", "MA_PSU"):
                attrs["unit_label"] = attrs["rank_text"] or ln["label"]
                attrs["rank"] = None
            elif ln["rank"]:
                attrs["rank"] = ln["rank"]
            else:
                attrs["rank"] = C.RANK_BY_LABEL.get(attrs["rank_text"]) if attrs["rank_text"] else None
            if res["row_type"] == "ADDITIONAL":
                attrs["supply"] = {"Force Supplied": "FS", "Mutual Aid Requested": "MA"}.get(supply_txt)
            else:
                attrs["supply"] = res["supply"]
            sub.lines[rk] = attrs
            for c, (risk, day, meas) in inputs.items():
                val, state = _num(ws.cell(r, c).value)
                sub.cells.append(dict(row_key=rk, risk=risk, day=day, measure=meas, value=val, state=state,
                                      raw=ws.cell(r, c).value, cell=ws.cell(r, c).coordinate, sheet=ws.title))
        for missing in sorted(self.expected_keys[coll] - seen):
            sub.structure_issues.append((coll, missing, None, "DQ03", SEV_ERROR, "Expected row missing"))

    # ---- calculation ---------------------------------------------------------
    def calculate(self, submissions, fixtures=None):
        subs = self._latest_per_force(submissions)
        dq, inputs = [], []
        for s in submissions:
            if s not in subs:
                dq.append(self._dq(s, None, None, "DQ16", SEV_INFO,
                                   "Superseded by a later submission for the same force"))
        for s in subs:
            # structure issue tuple: (collection, row_key, cell, code, severity, detail)
            dq.extend(self._dq(s, i[0], i[1], i[3], i[4], i[5], i[2]) for i in s.structure_issues)
            dq.extend(self._check_cover(s))
            inputs.extend(self._clean_inputs(s, dq))
        deployed, packages = self._deploy(inputs, dq)
        self._cost(deployed, dq)
        wt_total = self._wt_tournament(deployed, packages)
        cal = self.expand_calendar(deployed, packages, fixtures) if fixtures else []
        return dict(inputs=inputs, deployed=deployed, packages=packages, dq=dq,
                    wt_tournament=wt_total, calendar=cal)

    @staticmethod
    def _latest_per_force(subs):
        best = {}
        for s in subs:
            f = s.cover.get("Force") or s.file
            n = s.cover.get("SubmissionNo") or 0
            if f not in best or n > (best[f].cover.get("SubmissionNo") or 0):
                best[f] = s
        return list(best.values())

    @staticmethod
    def _dq(sub, coll, row_key, code, sev, detail, cell=None):
        return dict(submission=sub.submission_id, force=sub.cover.get("Force"), file=sub.file,
                    collection=coll, row_key=row_key, cell=cell, code=code, severity=sev, detail=detail)

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
        """Apply blank/zero/N-A/invalid rules; return one record per line x risk x day."""
        grouped = defaultdict(dict)
        for c in s.cells:
            ln = s.lines[c["row_key"]]
            res = ln["res"]
            sec = f"{res['collection']}-{res['section_code']}"
            status = s.section_status.get(sec, "Not started")
            applicable = C.applicable(res, c["measure"], c["day"] if res["collection"] == "MD" else None)
            if res["row_type"] == "ADDITIONAL" and c["measure"] == "TOD" and ln["supply"] == "MA" and not self.ma_tod:
                applicable = False
            if res["supply"] == "MA" and c["measure"] == "TOD" and self.ma_tod:
                applicable = True
            val, state = c["value"], c["state"]
            if state == "INVALID":
                dq.append(self._dq(s, res["collection"], c["row_key"], "DQ04", SEV_ERROR,
                                   f"Non-numeric value '{c['raw']}'", c["cell"]))
                val = None
            if not applicable:
                if state == "VALUE":
                    dq.append(self._dq(s, res["collection"], c["row_key"], "DQ06", SEV_WARN,
                                       "Value in a not-applicable cell - excluded", c["cell"]))
                val, state = None, "NA"
            elif val is not None:
                if val < 0:
                    dq.append(self._dq(s, res["collection"], c["row_key"], "DQ04", SEV_ERROR,
                                       "Negative value - excluded", c["cell"]))
                    val, state = None, "INVALID"
                elif c["measure"] == "NUM" and val != int(val):
                    dq.append(self._dq(s, res["collection"], c["row_key"], "DQ04", SEV_ERROR,
                                       "Number required must be a whole number - excluded", c["cell"]))
                    val, state = None, "INVALID"
                elif c["measure"] in ("TOD", "XH") and val > 24:
                    dq.append(self._dq(s, res["collection"], c["row_key"], "DQ05", SEV_ERROR,
                                       "Hours above 24 - excluded", c["cell"]))
                    val, state = None, "INVALID"
            elif state == "BLANK" and status in ("Complete", "Not required"):
                val, state = 0.0, "BLANK_AS_ZERO"
            key = (c["row_key"], c["risk"], c["day"])
            grouped[key][c["measure"]] = (val, state)
        out = []
        for (rk, risk, day), m in grouped.items():
            ln = s.lines[rk]
            res = ln["res"]
            rec = dict(submission=s.submission_id, force=s.cover.get("Force"), collection=res["collection"],
                       row_key=rk, resource_id=res["resource_id"], line=ln["line"], row_type=res["row_type"],
                       section=f"{res['collection']}-{res['section_code']}", portfolio=res["portfolio"],
                       resource=res["name"], supply=ln["supply"], rank=ln["rank"], role=ln["role"],
                       location=ln["location"] if res["location"] == "Option" else res["location"],
                       unit_label=ln.get("unit_label"), risk=risk, day=day,
                       num=m.get("NUM", (None, "NA"))[0], tod=m.get("TOD", (None, "NA"))[0],
                       xh=m.get("XH", (None, "NA"))[0],
                       num_state=m.get("NUM", (None, "NA"))[1],
                       status=s.section_status.get(f"{res['collection']}-{res['section_code']}", "Not started"))
            self._row_checks(s, rec, dq)
            out.append(rec)
        for sec, st in s.section_status.items():
            if st not in ("Complete", "Not required"):
                dq.append(self._dq(s, sec[:2], f"SECTION|{sec}", "DQ12", SEV_WARN,
                                   f"Section '{C.SECTIONS.get((sec[:2], sec[3:]), sec)}' status: {st}"))
        return out

    def _row_checks(self, s, rec, dq):
        rt = rec["row_type"]
        has_num = (rec["num"] or 0) > 0
        if rt in OFFICER_TYPES and has_num:
            if rec["tod"] in (None, 0) and rec["xh"] in (None, 0):
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ07", SEV_WARN,
                                   f"{rec['risk']} {rec['day']}: number entered but no hours"))
            if (rec["tod"] or 0) + (rec["xh"] or 0) > 24:
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ05", SEV_ERROR,
                                   f"{rec['risk']} {rec['day']}: Normal ToD + Extra Hours above 24"))
            if not rec["rank"]:
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ09", SEV_ERROR,
                                   f"{rec['risk']} {rec['day']}: rank not selected - not costed"))
            if rt == "ROLE" and not rec["role"]:
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ10", SEV_WARN,
                                   f"{rec['risk']} {rec['day']}: role profile not selected"))
            if rt == "ADDITIONAL" and not rec["supply"]:
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ11", SEV_ERROR,
                                   f"{rec['risk']} {rec['day']}: supply type not selected - not costed"))
            if rec["location"] in (None, "", "Option"):
                dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ15", SEV_WARN,
                                   f"{rec['risk']} {rec['day']}: location option not selected"))
        if rt in OFFICER_TYPES and not has_num and ((rec["tod"] or 0) > 0 or (rec["xh"] or 0) > 0):
            dq.append(self._dq(s, rec["collection"], rec["row_key"], "DQ08", SEV_WARN,
                               f"{rec['risk']} {rec['day']}: hours entered but no number"))

    # ---- deployment ----------------------------------------------------------
    def _deploy(self, inputs, dq):
        by_key = {(i["submission"], i["resource_id"], i["line"], i["risk"], i["day"]): i for i in inputs}
        deployed, packages = [], []
        md1_factor = float(self.params["MD1_PSU_Factor"])
        for i in inputs:
            rt = i["row_type"]
            base = {k: i[k] for k in ("submission", "force", "collection", "resource_id", "resource", "portfolio",
                                      "section", "supply", "role", "location", "risk", "day")}
            if rt in OFFICER_TYPES:
                n = i["num"] or 0
                if n <= 0:
                    continue
                deployed.append(dict(base, rank=i["rank"], source="Direct", units=None, officers=n,
                                     tod_hours=n * (i["tod"] or 0), xh_hours=n * (i["xh"] or 0)))
            elif rt == "PSU_COUNT":
                units = (i["num"] or 0) * (md1_factor if i["day"] == "MD+1" else 1)
                if units <= 0:
                    continue
                res = C.RESOURCE_BY_ID[i["resource_id"]]
                lvl = res["units"][i["line"] - 1]
                cfg = C.RESOURCE_BY_ID[res["psu_config"][lvl]]
                found = False
                for cl in C.lines(cfg):
                    c = by_key.get((i["submission"], cfg["resource_id"], cl["line"], i["risk"], i["day"]))
                    per = (c or {}).get("num") or 0
                    if per <= 0:
                        continue
                    found = True
                    off = units * per
                    deployed.append(dict(base, rank=cl["rank"], source="PSU expansion", units=units,
                                         unit_label=i["unit_label"], officers=off,
                                         tod_hours=off * (c["tod"] or 0), xh_hours=off * (c["xh"] or 0)))
                if not found:
                    dq.append(dict(submission=i["submission"], force=i["force"], file=None,
                                   collection=i["collection"], row_key=i["row_key"], cell=None, code="DQ13",
                                   severity=SEV_ERROR,
                                   detail=f"{i['risk']} {i['day']}: {units:g} units entered but PSU configuration "
                                          f"'{cfg['name']}' is empty for this risk/day"))
            elif rt == "MA_PSU":
                pass
        # mutual-aid PSU packages
        grp = defaultdict(list)
        for i in inputs:
            if i["row_type"] == "MA_PSU" and (i["num"] or 0) > 0:
                grp[(i["submission"], i["resource_id"], i["line"], i["risk"])].append(i)
        for (sub, rid, line, risk), rows in grp.items():
            res = C.RESOURCE_BY_ID[rid]
            if res["collection"] == "WT":
                basis = self.params["MAPSU_Basis_WT"]
            elif res["matchday_only"]:
                basis = self.params["MAPSU_Basis_Matchday"]
            else:
                basis = self.params["MAPSU_Basis_NonMatchday"]
            if basis.endswith("MAXDAY"):
                n_pk = max(r["num"] for r in rows)
            elif basis == "PER_UNIT_PER_FIXTURE":
                n_pk = sum(r["num"] for r in rows if r["day"] == "MD")
            else:  # PER_DAY (legacy behaviour / WT per day)
                n_pk = sum(r["num"] for r in rows)
            r0 = rows[0]
            unit = r0["unit_label"]
            rate = self.rates.package(unit)
            if rate is None:
                dq.append(dict(submission=sub, force=r0["force"], file=None, collection=r0["collection"],
                               row_key=r0["row_key"], cell=None, code="DQ14", severity=SEV_WARN,
                               detail=f"No mutual-aid package rate for '{unit}' - cost not calculated (OI-03)"))
            packages.append(dict(submission=sub, force=r0["force"], collection=r0["collection"], resource_id=rid,
                                 resource=res["name"], portfolio=res["portfolio"], location=res["location"],
                                 risk=risk, unit_label=unit, basis=basis, packages=n_pk,
                                 package_rate=rate, cost=(n_pk * rate) if rate is not None else None,
                                 days={r["day"]: r["num"] for r in rows}))
        return deployed, packages

    def _cost(self, deployed, dq):
        for d in deployed:
            rates = self.rates.hourly(d["supply"], d["rank"]) if d["rank"] and d["supply"] else None
            if rates is None:
                d.update(tod_rate=None, ot_rate=None, tod_cost=None, xh_cost=None, cost=None)
                if d["rank"] and d["supply"]:
                    dq.append(dict(submission=d["submission"], force=d["force"], file=None,
                                   collection=d["collection"], row_key=d["resource_id"], cell=None, code="DQ14",
                                   severity=SEV_WARN,
                                   detail=f"No {d['supply']} rate for rank {d['rank']} - not costed"))
                continue
            d["tod_rate"], d["ot_rate"] = rates
            d["tod_cost"] = d["tod_hours"] * rates[0]
            d["xh_cost"] = d["xh_hours"] * rates[1]
            d["cost"] = d["tod_cost"] + d["xh_cost"]

    def _wt_tournament(self, deployed, packages):
        days = {r: n for r, n, _ in R.WT_DAY_ASSUMPTIONS}
        out = defaultdict(float)
        for d in deployed:
            if d["collection"] == "WT" and d["cost"] is not None:
                out[(d["force"], d["risk"])] += d["cost"] * days[d["risk"]]
        for p in packages:
            if p["collection"] == "WT" and p["cost"] is not None:
                out[(p["force"], p["risk"])] += p["cost"] * days[p["risk"]]
        return dict(out)

    # ---- calendar ------------------------------------------------------------
    @staticmethod
    def expand_calendar(deployed, packages, fixtures):
        """
        fixtures: list of dicts(fixture_id, date, host_force, risk).
        Each MD demand row is placed on fixture date + day offset. Where dates
        coincide (D-02) rows are simply ADDED - no precedence rules.
        """
        offset = {d: o for d, _, o, _ in C.MD_DAYS}
        cal = []
        for fx in fixtures:
            for d in deployed:
                if d["collection"] != "MD" or d["force"] != fx["host_force"] or d["risk"] != fx["risk"]:
                    continue
                cal.append(dict(d, fixture_id=fx["fixture_id"],
                                date=fx["date"] + dt.timedelta(days=offset[d["day"]])))
        return cal


def summarise(rows, keys, value="cost"):
    out = defaultdict(float)
    for r in rows:
        if r.get(value) is not None:
            out[tuple(r[k] for k in keys)] += r[value]
    return dict(out)
