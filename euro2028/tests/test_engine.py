"""
Unit tests for the EURO 2028 template and calculation rules (both rule sets).

Run:  python3 -m unittest discover -s euro2028/tests -v

Expected values are hand calculations using the LEGACY workbook's cached
June-2028 rates (Baseline Costs column X), so the tests also prove that the
rate derivation reproduces the old workbook.
"""
import datetime as dt
import itertools
import os
import sys
import tempfile
import unittest

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.join(HERE, "..", "build"), os.path.join(HERE, "..", "engine"), HERE]

import build_template  # noqa: E402
import catalogue as C  # noqa: E402
import reference_data as RD  # noqa: E402
from e28_engine import Engine, Rates, default_params, summarise  # noqa: E402
from fill import fill, sheet_maps  # noqa: E402

LEGACY = {
    ("MA", "CSUPT"): 106.6548680856368, ("MA", "SUPT"): 92.3808239153443, ("MA", "CI"): 74.90307421259567,
    ("MA", "INSP"): 67.39850371876011, ("MA", "SGT"): 67.4643332844955, ("MA", "PC"): 54.22161897739092,
    ("HF", "CSUPT"): 99.825336, ("HF", "SUPT"): 85.778316, ("HF", "CI"): 70.93476,
    ("HF", "INSP"): 67.178124, ("HF", "SGT"): 55.208556, ("HF", "PC"): 45.973044,
}
PKG = 35716.1613354
HF_TOD_SGT, HF_TOD_PC = 61.49 * 1.04 * 1.035, 49.42 * 1.04 * 1.035
FORCE = "SYNTHETIC TEST FORCE"
_COUNTER = itertools.count(1)
COVER = {"cv_Force": FORCE, "cv_ContactName": "Test", "cv_ContactEmail": "t@example.invalid", "cv_SubmissionNo": 1}


class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.template = build_template.build(os.path.join(cls.tmp.name, "template.xlsx"))
        cls.leg = Engine()                                   # default RuleSet = LEGACY
        cls.agr = Engine(params={"RuleSet": "AGREED"})

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def sub(self, **kw):
        kw.setdefault("cover", COVER)
        return fill(self.template, os.path.join(self.tmp.name, f"s{next(_COUNTER)}.xlsx"), **kw)

    def run_both(self, fixtures=None, **kw):
        p = self.sub(**kw)
        return (self.leg.calculate([self.leg.read_submission(p)], fixtures),
                self.agr.calculate([self.agr.read_submission(p)], fixtures))


def psu_cfg(rid, risk, day, comp=(1, 3, 18), tod=10, xh=4):
    v = {}
    for line, n in enumerate(comp, 1):
        v.update({("MD", f"{rid}|{line}", f"{risk}|{day}|NUM"): n, ("MD", f"{rid}|{line}", f"{risk}|{day}|TOD"): tod,
                  ("MD", f"{rid}|{line}", f"{risk}|{day}|XH"): xh})
    return v


class TestRates(Base):
    def test_rates_reproduce_legacy(self):
        r = Rates(default_params())
        for (table, rank), legacy in LEGACY.items():
            self.assertAlmostEqual(r.table_rate(table, rank)[1], legacy, places=6, msg=f"{table} {rank}")
        self.assertAlmostEqual(r.package("Level 2 PSU"), PKG, places=4)

    def test_default_ruleset_is_legacy(self):
        self.assertEqual(default_params()["RuleSet"], "LEGACY")

    def test_legacy_extra_hours_only_and_zero_ranks(self):
        r = Rates(default_params())
        self.assertEqual(r.hourly("FS", "SGT")[0], 0.0)                     # no ToD rate in legacy
        self.assertAlmostEqual(r.hourly("FS", "SGT")[1], LEGACY[("HF", "SGT")], places=6)
        for rk in ("INSP", "CI", "SUPT", "CSUPT", "ACC", "DI"):
            self.assertEqual(r.hourly("FS", rk), (0.0, 0.0), rk)
        self.assertAlmostEqual(r.hourly("MA", "INSP")[1], LEGACY[("MA", "INSP")], places=6)
        self.assertIsNone(r.package("Level 2 Serial"))


class TestTemplate(Base):
    def test_layout_matches_forces_form(self):
        wb = openpyxl.load_workbook(self.template)
        self.assertEqual(wb.sheetnames[:4], ["Cover", "Guidance", "Match Day Resource", "Wider Tournament Resource"])
        ws = wb["Match Day Resource"]
        self.assertEqual([ws.cell(5, c).value for c in (1, 2, 3)],
                         ["Cell Input Instructions", "Resource", "Rank (Drop Down)"])
        cm, rm = sheet_maps(ws)
        self.assertEqual(cm["L|MD-2|NUM"], 4)                                # inputs start in column D
        self.assertEqual(cm["SH|MD+1|XH"], 4 + 47)
        keys = [k for k in rm if not k.startswith("SECTION")]
        expected = [C.row_key(r["resource_id"], ln["line"]) for r in C.RESOURCES
                    if r["collection"] == "MD" for ln in C.lines(r)]
        self.assertEqual(keys, expected)
        self.assertTrue(ws.column_dimensions[openpyxl.utils.get_column_letter(cm["ROWKEY"])].hidden)

    def test_na_cells(self):
        ws = openpyxl.load_workbook(self.template)["Match Day Resource"]
        cm, rm = sheet_maps(ws)
        locked = lambda rk, ck: ws.cell(rm[rk], cm[ck]).protection.locked  # noqa: E731
        self.assertTrue(locked("MD-PSC-07|1", "H|MD-2|NUM"))                 # matchday-only
        self.assertFalse(locked("MD-PSC-07|1", "H|MD|NUM"))
        self.assertTrue(locked("MD-PSC-01|1", "L|MD|NUM"))                   # non-matchday: MD N/A
        self.assertFalse(locked("MD-PSC-01|1", "L|MD-1|NUM"))
        self.assertTrue(locked("MD-EGT-02|1", "M|MD|TOD"))                   # MA ToD
        self.assertTrue(locked("MD-POP-03|1", "M|MD|XH"))                    # count rows: hours N/A


class TestLegacyCosting(Base):
    def test_in_stadium_psu(self):
        v = psu_cfg("MD-PSC-07", "H", "MD")
        v[("MD", "MD-POP-07|1", "H|MD|NUM")] = 6
        leg, agr = self.run_both(values=v)
        rows = {d["rank"]: d for d in leg["deployed"] if d["source"] == "PSU expansion"}
        self.assertEqual({k: d["officers"] for k, d in rows.items()}, {"INSP": 6, "SGT": 18, "PC": 108})
        self.assertAlmostEqual(rows["INSP"]["cost"], 6 * 4 * LEGACY[("MA", "INSP")], places=4)
        self.assertAlmostEqual(rows["SGT"]["cost"], 18 * 4 * LEGACY[("HF", "SGT")], places=4)
        self.assertAlmostEqual(rows["PC"]["cost"], 108 * 4 * LEGACY[("HF", "PC")], places=4)
        arows = {d["rank"]: d for d in agr["deployed"] if d["source"] == "PSU expansion"}
        self.assertEqual(arows["INSP"]["cost"], 0)
        self.assertAlmostEqual(arows["PC"]["cost"], 108 * (10 * HF_TOD_PC + 4 * LEGACY[("HF", "PC")]), places=4)

    def test_tod_not_costed_in_legacy(self):
        v = {("MD", "MD-EGT-01|1", "L|MD|NUM"): 2, ("MD", "MD-EGT-01|1", "L|MD|TOD"): 10,
             ("MD", "MD-EGT-01|1", "L|MD|XH"): 2}
        leg, agr = self.run_both(values=v)
        self.assertAlmostEqual(leg["deployed"][0]["cost"], 2 * 2 * LEGACY[("HF", "SGT")], places=6)
        self.assertAlmostEqual(agr["deployed"][0]["cost"], 2 * (10 * HF_TOD_SGT + 2 * LEGACY[("HF", "SGT")]), places=6)

    def test_ma_package_basis(self):
        v = {("MD", "MD-POP-02|1", "L|MD-2|NUM"): 1, ("MD", "MD-POP-02|1", "L|MD-1|NUM"): 2,
             ("MD", "MD-POP-02|1", "L|MD+1|NUM"): 1, ("MD", "MD-POP-08|1", "H|MD|NUM"): 2}
        leg, agr = self.run_both(values=v)
        tot = lambda res, rid: sum(p["packages"] for p in res["packages"] if p["resource_id"] == rid)  # noqa: E731
        self.assertEqual(tot(leg, "MD-POP-02"), 4)      # legacy: per unit per day
        self.assertEqual(tot(agr, "MD-POP-02"), 2)      # agreed default: peak day
        self.assertEqual(tot(leg, "MD-POP-08"), 2)
        self.assertAlmostEqual(sum(p["cost"] for p in leg["packages"]), 6 * PKG, places=3)

    def test_additional_resources(self):
        v = {("MD", "MD-ADD-02|1", "L|MD|NUM"): 3, ("MD", "MD-ADD-02|1", "L|MD|XH"): 2}
        a = {("MD", "MD-ADD-02|1", "COLB"): "Hotel liaison", ("MD", "MD-ADD-02|1", "COLC"): "Constable"}
        leg, agr = self.run_both(values=v, attrs=a)
        self.assertAlmostEqual(leg["deployed"][0]["cost"], 6 * LEGACY[("MA", "SGT")], places=6)   # legacy fixed rate
        self.assertAlmostEqual(agr["deployed"][0]["cost"], 6 * LEGACY[("MA", "PC")], places=6)    # by rank

    def test_role_rows_by_selected_rank(self):
        v = {("MD", "MD-AAP-02|1", "H|MD|NUM"): 1, ("MD", "MD-AAP-02|1", "H|MD|XH"): 5,
             ("MD", "MD-AAP-01|1", "H|MD|NUM"): 2, ("MD", "MD-AAP-01|1", "H|MD|XH"): 5}
        a = {("MD", "MD-AAP-02|1", "COLA"): "Superintendent", ("MD", "MD-AAP-02|1", "COLC"): "CTSFO",
             ("MD", "MD-AAP-01|1", "COLA"): "Chief Inspector", ("MD", "MD-AAP-01|1", "COLC"): "ARV"}
        leg, _ = self.run_both(values=v, attrs=a)
        c = {d["resource_id"]: d["cost"] for d in leg["deployed"]}
        self.assertAlmostEqual(c["MD-AAP-02"], 5 * LEGACY[("MA", "SUPT")], places=6)
        self.assertEqual(c["MD-AAP-01"], 0)

    def test_wider_tournament(self):
        v = {("WT", "WT-POP-05|2", "H|WT|NUM"): 10, ("WT", "WT-POP-05|2", "H|WT|TOD"): 8,
             ("WT", "WT-POP-05|2", "H|WT|XH"): 2, ("WT", "WT-POP-02|1", "L|WT|NUM"): 1}
        leg, agr = self.run_both(values=v)
        self.assertAlmostEqual(leg["wt_tournament"][(FORCE, "H")], 10 * 2 * LEGACY[("HF", "SGT")] * 11, places=4)
        self.assertAlmostEqual(leg["wt_tournament"][(FORCE, "L")], PKG * 7, places=3)
        self.assertAlmostEqual(agr["wt_tournament"][(FORCE, "H")],
                               10 * (8 * HF_TOD_SGT + 2 * LEGACY[("HF", "SGT")]) * 11, places=4)


class TestCalendar(Base):
    FX = [dict(fixture_id="F1", date=dt.date(2028, 6, 12), host_force=FORCE, risk="H"),
          dict(fixture_id="F2", date=dt.date(2028, 6, 15), host_force=FORCE, risk="H")]

    def test_overlap_legacy_first_match_vs_agreed_add(self):
        v = {("MD", "MD-CTL-01|3", "H|MD+1|NUM"): 5, ("MD", "MD-CTL-01|3", "H|MD+1|XH"): 2,
             ("MD", "MD-CTL-01|3", "H|MD-2|NUM"): 7, ("MD", "MD-CTL-01|3", "H|MD-2|XH"): 2}
        leg, agr = self.run_both(fixtures=self.FX, values=v)
        L = summarise(leg["calendar"], ["date"], "officers")
        A = summarise(agr["calendar"], ["date"], "officers")
        d13 = (dt.date(2028, 6, 13),)
        self.assertEqual(L[d13], 7)      # legacy STD priority: MD-2 beats MD+1
        self.assertEqual(A[d13], 12)     # agreed: added
        self.assertEqual(L[(dt.date(2028, 6, 10),)], 7)
        self.assertEqual(L[(dt.date(2028, 6, 16),)], 5)

    def test_legacy_nonmatchday_priority(self):
        # Non-matchday PSU: MD-1 > MD-2 > MD+1. 14 Jun = MD-1 of F2 (15 Jun); 13 Jun = MD+1 of F1 and MD-2 of F2.
        v = {}
        for day, n in (("MD-2", 1), ("MD-1", 2), ("MD+1", 3)):
            v.update(psu_cfg("MD-PSC-01", "H", day, comp=(0, 1, 7), tod=10, xh=2))
            v[("MD", "MD-POP-01|1", f"H|{day}|NUM")] = n
        leg, _ = self.run_both(fixtures=self.FX, values=v)
        U = summarise(leg["calendar"], ["date"], "officers")
        self.assertEqual(U[(dt.date(2028, 6, 13),)], 1 * 8)    # MD-2 (1 unit x 8 officers), MD+1 ignored
        self.assertEqual(U[(dt.date(2028, 6, 14),)], 2 * 8)
        self.assertNotIn((dt.date(2028, 6, 12),), U)           # match day itself: not used for non-matchday
        self.assertEqual(U[(dt.date(2028, 6, 16),)], 3 * 8)

    def test_different_risks_are_added(self):
        fx = [self.FX[0], dict(self.FX[1], risk="M")]
        v = {("MD", "MD-CTL-01|3", "H|MD+1|NUM"): 5, ("MD", "MD-CTL-01|3", "H|MD+1|XH"): 1,
             ("MD", "MD-CTL-01|3", "M|MD-2|NUM"): 4, ("MD", "MD-CTL-01|3", "M|MD-2|XH"): 1}
        leg, _ = self.run_both(fixtures=fx, values=v)
        self.assertEqual(summarise(leg["calendar"], ["date"], "officers")[(dt.date(2028, 6, 13),)], 9)


class TestDataQuality(Base):
    def test_invalid_and_missing(self):
        v = {("MD", "MD-CUS-01|1", "L|MD|NUM"): 1.5, ("MD", "MD-CUS-01|2", "L|MD|NUM"): "abc",
             ("MD", "MD-CUS-01|1", "M|MD|NUM"): 2, ("MD", "MD-CUS-01|1", "M|MD|XH"): 30,
             ("MD", "MD-EGT-02|1", "M|MD|NUM"): 1, ("MD", "MD-EGT-02|1", "M|MD|TOD"): 8,
             ("MD", "MD-CBR-03|1", "H|MD|NUM"): 2, ("MD", "MD-CBR-03|1", "H|MD|XH"): 3,
             ("MD", "MD-POP-05|2", "L|MD|NUM"): 3}
        leg, _ = self.run_both(values=v)
        codes = [d["code"] for d in leg["dq"]]
        for code in ("DQ04", "DQ05", "DQ06", "DQ07", "DQ09", "DQ10", "DQ12", "DQ13"):
            self.assertIn(code, codes)

    def test_blank_vs_zero(self):
        leg, _ = self.run_both(status={"MD-CTL": "Complete"}, values={("MD", "MD-CTL-01|2", "L|MD|NUM"): 0})
        st = {i["line"]: i["num_state"] for i in leg["inputs"]
              if i["resource_id"] == "MD-CTL-01" and i["risk"] == "L" and i["day"] == "MD"}
        self.assertEqual((st[1], st[2]), ("BLANK_AS_ZERO", "VALUE"))
        self.assertFalse(any("Control Room" in d["detail"] for d in leg["dq"] if d["code"] == "DQ12"))

    def test_latest_submission_and_cover(self):
        p1 = self.sub(values={("MD", "MD-CTL-01|3", "L|MD|NUM"): 1})
        p2 = self.sub(cover=dict(COVER, cv_SubmissionNo=2), values={("MD", "MD-CTL-01|3", "L|MD|NUM"): 9})
        p3 = self.sub(cover={"cv_SubmissionNo": 1})
        res = self.leg.calculate([self.leg.read_submission(p) for p in (p1, p2, p3)])
        self.assertEqual(sum(d["officers"] for d in res["deployed"] if d["force"] == FORCE), 9)
        codes = {d["code"] for d in res["dq"]}
        self.assertTrue({"DQ16", "DQ01"} <= codes)


class TestOtherPages(unittest.TestCase):
    def test_command_team_matches_legacy_method(self):
        from e28_engine import command_team_costs
        rows = [dict(rank="CI", posts=1, start=dt.date(2028, 1, 1), end=dt.date(2028, 7, 1)),
                dict(rank="SGT", posts=1, start=dt.date(2027, 4, 1), end=dt.date(2028, 7, 1)),
                dict(rank="PC", posts=1, start=dt.date(2027, 4, 1), end=dt.date(2028, 7, 1))]
        out = command_team_costs(rows)
        sgt_first = next(x["cost"] for x in out if x["rank"] == "SGT")
        self.assertAlmostEqual(sgt_first, 8607.401229951885, places=6)        # legacy CTP!B17
        self.assertEqual(len([x for x in out if x["rank"] == "SGT"]), 16)      # Apr-27..Jul-28 inclusive

    def test_other_resources_rates(self):
        from e28_engine import other_resources_costs
        out = other_resources_costs([dict(category="Intel", rank="PC", hours_2728=24, hours_2829=280)])[0]
        self.assertAlmostEqual(out["cost_2728"], 1066.0416, places=4)           # legacy Other Resources!E17
        self.assertAlmostEqual(out["cost_2829"], 12872.45232, places=4)         # legacy Other Resources!G17


class TestMapping(unittest.TestCase):
    def test_mapping_complete(self):
        import legacy_mapping as M
        ids = {r["resource_id"] for r in C.RESOURCES}
        mapped = {n for m in M.MAPPING for n in m[3]}
        self.assertTrue(mapped <= ids, mapped - ids)
        self.assertTrue(set(M.ADDED) <= ids, set(M.ADDED) - ids)
        self.assertEqual(ids - mapped - set(M.ADDED), set())
        for m in M.MAPPING:
            self.assertEqual(bool(m[3]), m[4] != "Retired", m[2])

    def test_parameters_have_ruleset(self):
        self.assertIn("RuleSet", {p[0] for p in RD.PARAMETERS})


if __name__ == "__main__":
    unittest.main()
