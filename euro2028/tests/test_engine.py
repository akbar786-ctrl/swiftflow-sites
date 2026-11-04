"""
Unit tests for the EURO 2028 template and calculation rules.

Run:  python3 -m unittest discover -s euro2028/tests -v

Expected values are calculated by hand from the LEGACY workbook's cached
June-2028 rates (Baseline Costs column X), so these tests also prove that
the new rate derivation reproduces the old workbook.
"""
import datetime as dt
import os
import sys
import tempfile
import unittest

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "build"))
sys.path.insert(0, os.path.join(HERE, "..", "engine"))
sys.path.insert(0, HERE)

import build_template  # noqa: E402
import catalogue as C  # noqa: E402
from e28_engine import Engine, Rates, summarise  # noqa: E402
from fill import fill  # noqa: E402

# Legacy cached values (Baseline Costs!X10:X17, X27:X32)
LEGACY = {
    ("MA", "CSUPT"): 106.6548680856368, ("MA", "SUPT"): 92.3808239153443,
    ("MA", "CI"): 74.90307421259567, ("MA", "INSP"): 67.39850371876011,
    ("MA", "SGT"): 67.4643332844955, ("MA", "PC"): 54.22161897739092,
    ("HF", "CSUPT"): 99.825336, ("HF", "SUPT"): 85.778316, ("HF", "CI"): 70.93476,
    ("HF", "INSP"): 67.178124, ("HF", "SGT"): 55.208556, ("HF", "PC"): 45.973044,
}
LEGACY_PACKAGE = 35716.1613354
HF_TOD_SGT = 61.49 * 1.04 * 1.035     # Hourly Employable Cost (Baseline Costs!D22) uplifted
HF_TOD_PC = 49.42 * 1.04 * 1.035

COVER = {"cv_Force": "SYNTHETIC TEST FORCE", "cv_ContactName": "Test", "cv_ContactEmail": "t@example.org",
         "cv_SubmissionNo": 1, "cv_SubmissionDate": dt.datetime(2026, 11, 4)}


class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.template = build_template.build(os.path.join(cls.tmp.name, "template.xlsx"))
        cls.eng = Engine()

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def run_sub(self, name, engine=None, **kw):
        path = fill(self.template, os.path.join(self.tmp.name, name), **kw)
        eng = engine or self.eng
        return eng, eng.calculate([eng.read_submission(path)])


class TestRates(Base):
    def test_hourly_overtime_rates_reproduce_legacy(self):
        r = Rates()
        for (table, rank), legacy in LEGACY.items():
            base = next(b for b in __import__("reference_data").BASE_RATES if b[0] == table and b[1] == rank)
            factor = r.factor(table)
            self.assertAlmostEqual(base[3] * factor, legacy, places=6, msg=f"{table} {rank}")

    def test_package_reproduces_legacy(self):
        self.assertAlmostEqual(Rates().package("Level 2 PSU"), LEGACY_PACKAGE, places=4)

    def test_force_supplied_inspector_and_above_zero(self):
        r = Rates()
        for rk in ("INSP", "CI", "SUPT", "CSUPT", "ACC", "DI"):
            self.assertEqual(r.hourly("FS", rk), (0.0, 0.0), rk)
        self.assertAlmostEqual(r.hourly("FS", "SGT")[1], LEGACY[("HF", "SGT")], places=6)
        self.assertAlmostEqual(r.hourly("FS", "DS")[1], LEGACY[("HF", "SGT")], places=6)

    def test_mutual_aid_inspector_charged(self):
        self.assertAlmostEqual(Rates().hourly("MA", "INSP")[1], LEGACY[("MA", "INSP")], places=6)

    def test_missing_rate_returns_none(self):
        self.assertIsNone(Rates().package("Level 2 Serial"))


class TestTemplateStructure(Base):
    def test_every_row_key_once(self):
        wb = openpyxl.load_workbook(self.template)
        for coll in ("MD", "WT"):
            ws = wb[f"{coll} Collection"]
            keys = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
            keys = [k for k in keys if k and "|" in k and not k.startswith("SECTION")]
            expected = [C.row_key(r["resource_id"], ln["line"]) for r in C.RESOURCES
                        if r["collection"] == coll for ln in C.lines(r)]
            self.assertEqual(keys, expected)

    def test_column_keys(self):
        wb = openpyxl.load_workbook(self.template)
        md = [c.value for c in wb["MD Collection"][6] if c.value and c.value.count("|") == 2
              and not c.value.startswith("TOTAL")]
        wt = [c.value for c in wb["WT Collection"][6] if c.value and c.value.count("|") == 2
              and not c.value.startswith("TOTAL")]
        self.assertEqual(len(md), 4 * 4 * 3)    # risk x day x measure
        self.assertEqual(len(wt), 4 * 3)
        self.assertEqual(md[:4], ["L|MD-2|NUM", "L|MD-2|TOD", "L|MD-2|XH", "L|MD-1|NUM"])
        self.assertIn("SH|MD|NUM", md)

    def test_na_cells_locked_inputs_unlocked(self):
        wb = openpyxl.load_workbook(self.template)
        ws = wb["MD Collection"]
        cm = {ws.cell(6, c).value: c for c in range(1, ws.max_column + 1)}
        rm = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
        # matchday-only config: MD-2 locked, MD open
        self.assertTrue(ws.cell(rm["MD-PSC-07|1"], cm["H|MD-2|NUM"]).protection.locked)
        self.assertFalse(ws.cell(rm["MD-PSC-07|1"], cm["H|MD|NUM"]).protection.locked)
        # MA ToD locked, MA extra hours open
        self.assertTrue(ws.cell(rm["MD-EGT-02|1"], cm["M|MD|TOD"]).protection.locked)
        self.assertFalse(ws.cell(rm["MD-EGT-02|1"], cm["M|MD|XH"]).protection.locked)
        # PSU count rows: hours locked
        self.assertTrue(ws.cell(rm["MD-POP-01|1"], cm["L|MD-1|TOD"]).protection.locked)
        self.assertTrue(ws.protection.sheet)


class TestCalculations(Base):
    def test_psu_expansion_and_cost(self):
        v = {("MD", "MD-PSC-07|1", "H|MD|NUM"): 1, ("MD", "MD-PSC-07|1", "H|MD|TOD"): 10,
             ("MD", "MD-PSC-07|1", "H|MD|XH"): 4,
             ("MD", "MD-PSC-07|2", "H|MD|NUM"): 3, ("MD", "MD-PSC-07|2", "H|MD|TOD"): 10,
             ("MD", "MD-PSC-07|2", "H|MD|XH"): 4,
             ("MD", "MD-PSC-07|3", "H|MD|NUM"): 18, ("MD", "MD-PSC-07|3", "H|MD|TOD"): 10,
             ("MD", "MD-PSC-07|3", "H|MD|XH"): 4,
             ("MD", "MD-POP-07|1", "H|MD|NUM"): 6}
        _, res = self.run_sub("psu.xlsx", cover=COVER, values=v)
        rows = [d for d in res["deployed"] if d["source"] == "PSU expansion"]
        off = {d["rank"]: d["officers"] for d in rows}
        self.assertEqual(off, {"INSP": 6, "SGT": 18, "PC": 108})
        self.assertEqual(sum(d["tod_hours"] + d["xh_hours"] for d in rows), 1848)   # guidance worked example
        expected = 18 * (10 * HF_TOD_SGT + 4 * LEGACY[("HF", "SGT")]) + \
            108 * (10 * HF_TOD_PC + 4 * LEGACY[("HF", "PC")])
        self.assertAlmostEqual(sum(d["cost"] for d in rows), expected, places=4)
        self.assertEqual(next(d for d in rows if d["rank"] == "INSP")["cost"], 0)

    def test_psu_count_without_config_flags(self):
        _, res = self.run_sub("nocfg.xlsx", cover=COVER, values={("MD", "MD-POP-05|2", "L|MD|NUM"): 3})
        self.assertIn("DQ13", {d["code"] for d in res["dq"]})

    def test_ma_package_bases(self):
        v = {("MD", "MD-POP-08|1", "H|MD|NUM"): 2,      # matchday-only MA
             ("MD", "MD-POP-02|1", "L|MD-2|NUM"): 1, ("MD", "MD-POP-02|1", "L|MD-1|NUM"): 2,
             ("MD", "MD-POP-02|1", "L|MD|NUM"): 2, ("MD", "MD-POP-02|1", "L|MD+1|NUM"): 1}
        _, res = self.run_sub("pkg.xlsx", cover=COVER, values=v)
        pk = {p["resource_id"]: p for p in res["packages"]}
        self.assertAlmostEqual(pk["MD-POP-08"]["cost"], 2 * LEGACY_PACKAGE, places=3)
        self.assertEqual(pk["MD-POP-02"]["packages"], 2)                  # peak day
        legacy = Engine(params={"MAPSU_Basis_NonMatchday": "PER_UNIT_PER_DAY"})
        _, res2 = self.run_sub("pkg2.xlsx", engine=legacy, cover=COVER, values=v)
        self.assertEqual({p["resource_id"]: p for p in res2["packages"]}["MD-POP-02"]["packages"], 6)

    def test_serial_package_has_no_rate(self):
        _, res = self.run_sub("serial.xlsx", cover=COVER, values={("MD", "MD-POP-04|1", "M|MD|NUM"): 1},
                              attrs={("MD", "MD-POP-04|1", "RANK"): "Level 2 Serial"})
        self.assertIsNone(res["packages"][0]["cost"])
        self.assertIn("DQ14", {d["code"] for d in res["dq"]})

    def test_mutual_aid_officer_extra_hours_only(self):
        v = {("MD", "MD-EGT-02|1", "M|MD|NUM"): 4, ("MD", "MD-EGT-02|1", "M|MD|XH"): 6,
             ("MD", "MD-EGT-02|1", "M|MD|TOD"): 10}   # typed into a locked N/A cell
        _, res = self.run_sub("ma.xlsx", cover=COVER, values=v)
        d = next(x for x in res["deployed"] if x["resource_id"] == "MD-EGT-02")
        self.assertEqual(d["tod_hours"], 0)
        self.assertAlmostEqual(d["cost"], 4 * 6 * LEGACY[("MA", "SGT")], places=6)
        self.assertIn("DQ06", {x["code"] for x in res["dq"]})

    def test_ma_tod_switch(self):
        eng = Engine(params={"MA_ToD_Collected": "TRUE"})
        v = {("MD", "MD-EGT-02|1", "M|MD|NUM"): 4, ("MD", "MD-EGT-02|1", "M|MD|XH"): 6,
             ("MD", "MD-EGT-02|1", "M|MD|TOD"): 10}
        _, res = self.run_sub("ma2.xlsx", engine=eng, cover=COVER, values=v)
        d = next(x for x in res["deployed"] if x["resource_id"] == "MD-EGT-02")
        self.assertAlmostEqual(d["cost"], 4 * 16 * LEGACY[("MA", "SGT")], places=6)

    def test_role_rows_priced_by_selected_rank(self):
        v = {("MD", "MD-AAP-01|1", "H|MD|NUM"): 2, ("MD", "MD-AAP-01|1", "H|MD|XH"): 5,
             ("MD", "MD-AAP-02|1", "H|MD|NUM"): 1, ("MD", "MD-AAP-02|1", "H|MD|XH"): 5}
        a = {("MD", "MD-AAP-01|1", "RANK"): "Chief Inspector", ("MD", "MD-AAP-01|1", "ROLE"): "ARV",
             ("MD", "MD-AAP-02|1", "RANK"): "Superintendent", ("MD", "MD-AAP-02|1", "ROLE"): "CTSFO"}
        _, res = self.run_sub("role.xlsx", cover=COVER, values=v, attrs=a)
        c = {d["resource_id"]: d["cost"] for d in res["deployed"]}
        self.assertEqual(c["MD-AAP-01"], 0)                            # force-supplied CI
        self.assertAlmostEqual(c["MD-AAP-02"], 5 * LEGACY[("MA", "SUPT")], places=6)

    def test_additional_resource(self):
        v = {("MD", "MD-ADD-01|1", "L|MD|NUM"): 3, ("MD", "MD-ADD-01|1", "L|MD|XH"): 2,
             ("MD", "MD-ADD-01|2", "L|MD|NUM"): 1, ("MD", "MD-ADD-01|2", "L|MD|XH"): 2}
        a = {("MD", "MD-ADD-01|1", "ROLE"): "Hotel liaison", ("MD", "MD-ADD-01|1", "SUPPLY"): "Mutual Aid Requested",
             ("MD", "MD-ADD-01|1", "RANK"): "Constable", ("MD", "MD-ADD-01|2", "SUPPLY"): "Force Supplied"}
        _, res = self.run_sub("add.xlsx", cover=COVER, values=v, attrs=a)
        costs = [d for d in res["deployed"] if d["resource_id"] == "MD-ADD-01"]
        self.assertAlmostEqual(next(d for d in costs if d["rank"] == "PC")["cost"],
                               3 * 2 * LEGACY[("MA", "PC")], places=6)
        self.assertIsNone(next(d for d in costs if d["rank"] is None)["cost"])
        self.assertIn("DQ09", {d["code"] for d in res["dq"]})

    def test_invalid_values(self):
        v = {("MD", "MD-CUS-01|1", "L|MD|NUM"): 1.5, ("MD", "MD-CUS-01|2", "L|MD|NUM"): "abc",
             ("MD", "MD-CUS-01|1", "M|MD|NUM"): 2, ("MD", "MD-CUS-01|1", "M|MD|XH"): 30}
        _, res = self.run_sub("bad.xlsx", cover=COVER, values=v)
        codes = [d["code"] for d in res["dq"]]
        self.assertGreaterEqual(codes.count("DQ04"), 2)
        self.assertIn("DQ05", codes)
        self.assertFalse([d for d in res["deployed"] if d["resource_id"] == "MD-CUS-01" and d["risk"] == "L"])

    def test_blank_zero_and_section_status(self):
        _, res = self.run_sub("status.xlsx", cover=COVER, status={"MD-CTL": "Complete"},
                              values={("MD", "MD-CTL-01|2", "L|MD|NUM"): 0})
        ctl = [i for i in res["inputs"] if i["resource_id"] == "MD-CTL-01" and i["risk"] == "L" and i["day"] == "MD"]
        states = {i["line"]: i["num_state"] for i in ctl}
        self.assertEqual(states[1], "BLANK_AS_ZERO")
        self.assertEqual(states[2], "VALUE")
        incomplete = [d for d in res["dq"] if d["code"] == "DQ12"]
        self.assertTrue(any("Custody" in d["detail"] for d in incomplete))
        self.assertFalse(any("Control Room" in d["detail"] for d in incomplete))

    def test_wider_tournament_days(self):
        v = {("WT", "WT-POP-05|2", "H|WT|NUM"): 10, ("WT", "WT-POP-05|2", "H|WT|TOD"): 8,
             ("WT", "WT-POP-05|2", "H|WT|XH"): 2, ("WT", "WT-POP-02|1", "L|WT|NUM"): 1}
        _, res = self.run_sub("wt.xlsx", cover=COVER, values=v)
        per_day = 10 * (8 * HF_TOD_SGT + 2 * LEGACY[("HF", "SGT")])
        self.assertAlmostEqual(res["wt_tournament"][("SYNTHETIC TEST FORCE", "H")], per_day * 11, places=4)
        self.assertAlmostEqual(res["wt_tournament"][("SYNTHETIC TEST FORCE", "L")], LEGACY_PACKAGE * 7, places=3)

    def test_calendar_overlap_adds(self):
        v = {("MD", "MD-CTL-01|3", "H|MD+1|NUM"): 5, ("MD", "MD-CTL-01|3", "H|MD+1|TOD"): 8,
             ("MD", "MD-CTL-01|3", "H|MD-2|NUM"): 7, ("MD", "MD-CTL-01|3", "H|MD-2|TOD"): 8}
        path = fill(self.template, os.path.join(self.tmp.name, "cal.xlsx"), cover=COVER, values=v)
        fx = [dict(fixture_id="F1", date=dt.date(2028, 6, 12), host_force="SYNTHETIC TEST FORCE", risk="H"),
              dict(fixture_id="F2", date=dt.date(2028, 6, 15), host_force="SYNTHETIC TEST FORCE", risk="H")]
        res = self.eng.calculate([self.eng.read_submission(path)], fixtures=fx)
        by_date = summarise(res["calendar"], ["date"], "officers")
        self.assertEqual(by_date[(dt.date(2028, 6, 13),)], 12)    # MD+1 of F1 (5) + MD-2 of F2 (7)
        self.assertEqual(by_date[(dt.date(2028, 6, 10),)], 7)
        self.assertEqual(by_date[(dt.date(2028, 6, 16),)], 5)

    def test_latest_submission_wins(self):
        p1 = fill(self.template, os.path.join(self.tmp.name, "s1.xlsx"), cover=COVER,
                  values={("MD", "MD-CTL-01|3", "L|MD|NUM"): 1})
        p2 = fill(self.template, os.path.join(self.tmp.name, "s2.xlsx"), cover=dict(COVER, cv_SubmissionNo=2),
                  values={("MD", "MD-CTL-01|3", "L|MD|NUM"): 9})
        res = self.eng.calculate([self.eng.read_submission(p1), self.eng.read_submission(p2)])
        self.assertEqual(sum(d["officers"] for d in res["deployed"]), 9)
        self.assertIn("DQ16", {d["code"] for d in res["dq"]})

    def test_missing_force_flagged(self):
        _, res = self.run_sub("noforce.xlsx", cover={"cv_SubmissionNo": 1})
        self.assertIn("DQ01", {d["code"] for d in res["dq"] if d["severity"] == "Error"})


class TestMapping(unittest.TestCase):
    def test_mapping_complete_and_valid(self):
        import legacy_mapping as M
        ids = {r["resource_id"] for r in C.RESOURCES}
        mapped = {n for *_, new, _, _ in [(m[0], m[1], m[2], m[3], m[4], m[5]) for m in M.MAPPING] for n in new}
        self.assertTrue(mapped <= ids, mapped - ids)
        self.assertTrue(set(M.ADDED) <= ids, set(M.ADDED) - ids)
        self.assertEqual(ids - mapped - set(M.ADDED), set(), "resources with no legacy mapping and not marked Added")
        for m in M.MAPPING:
            self.assertEqual(bool(m[3]), m[4] != "Retired", m[2])


if __name__ == "__main__":
    unittest.main()
