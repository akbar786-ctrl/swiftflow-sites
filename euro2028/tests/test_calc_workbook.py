"""
Reconciliation: backup calculation workbook (Excel formulas) vs the reference
engine (RuleSet = LEGACY), on the synthetic return including its deliberate errors.

The workbook is recalculated with LibreOffice (headless). Skipped if
LibreOffice is not installed.

Run:  python3 -m unittest euro2028.tests.test_calc_workbook  (or discover)
"""
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path[:0] = [os.path.join(HERE, "..", "build"), os.path.join(HERE, "..", "engine"), HERE]
import build_sample  # noqa: E402
import reference_data as R  # noqa: E402
from e28_engine import Engine  # noqa: E402

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")


def recalc(path, out_dir):
    subprocess.run([SOFFICE, "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", out_dir, path], check=True, capture_output=True, timeout=900)
    return os.path.join(out_dir, os.path.basename(path))


def sheet_by_key(ws, key_col=1, first_row=5, date_row=4, first_col=4, n_cols=None):
    out = {}
    for r in range(first_row, ws.max_row + 1):
        k = ws.cell(r, key_col).value
        if not k:
            continue
        for c in range(first_col, first_col + n_cols):
            d = ws.cell(date_row, c).value
            out[(k, d.date() if isinstance(d, dt.datetime) else d)] = ws.cell(r, c).value or 0
    return out


@unittest.skipUnless(SOFFICE, "LibreOffice not available")
class TestBackupWorkbookReconciles(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        sub, calc = build_sample.build(cls.tmp.name, with_errors=True)
        cls.wb = openpyxl.load_workbook(recalc(calc, os.path.join(cls.tmp.name, "recalc")), data_only=True)
        eng = Engine()
        assert eng.params["RuleSet"] == "LEGACY"
        fixtures = [dict(fixture_id=f"F{i}", date=d, host_force=build_sample.FORCE, risk=r)
                    for i, (d, r) in enumerate(build_sample.SYNTHETIC_FIXTURES, 1)]
        cls.res = eng.calculate([eng.read_submission(sub)], fixtures=fixtures)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def engine_calendar(self, value):
        out = defaultdict(float)
        for row in self.res["calendar"]:
            if row["kind"] == "officer":
                v = row["officers"] if value == "officers" else (row["cost"] or 0)
            else:
                v = row["units"] if value == "officers" else (row["cost"] or 0)
            out[(row["line_key"], row["date"])] += v
        return out

    def compare(self, sheet, value, tol):
        ws = self.wb[sheet]
        n_dates = sum(1 for c in range(4, ws.max_column + 1) if isinstance(ws.cell(4, c).value, (dt.date, dt.datetime)))
        wbv = sheet_by_key(ws, n_cols=n_dates)
        eng = self.engine_calendar(value)
        keys = set(wbv) | set(eng)
        diffs = [(k, wbv.get(k, 0), eng.get(k, 0)) for k in keys if abs((wbv.get(k, 0) or 0) - eng.get(k, 0)) > tol]
        self.assertFalse(diffs, f"{len(diffs)} differences, e.g. {sorted(diffs, key=str)[:5]}")
        self.assertGreater(sum(eng.values()), 0)
        return sum(eng.values())

    def test_md_officers_by_line_and_date(self):
        self.compare("MD Demand Template", "officers", 1e-9)

    def test_md_cost_by_line_and_date(self):
        self.compare("MD Cost Template", "cost", 0.005)

    def test_wt_cost(self):
        ws = self.wb["WT Cost Template"]
        days = {r: n for r, n, _ in R.WT_DAY_ASSUMPTIONS}
        eng = defaultdict(float)
        for row in self.res["deployed"] + self.res["packages"]:
            if row["collection"] == "WT" and row.get("cost") is not None:
                eng[row["line_key"]] += row["cost"] * days[row["risk"]]
        wbv = {ws.cell(r, 1).value: ws.cell(r, 12).value or 0 for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value}
        diffs = [(k, wbv.get(k, 0), eng.get(k, 0)) for k in set(wbv) | set(eng) if abs(wbv.get(k, 0) - eng.get(k, 0)) > 0.005]
        self.assertFalse(diffs, diffs[:5])
        self.assertAlmostEqual(sum(wbv.values()), sum(self.res["wt_tournament"].values()), places=2)

    def test_wt_officers_by_risk(self):
        ws = self.wb["WT Demand Template"]
        risks = [r for r, _, _ in R.WT_DAY_ASSUMPTIONS]
        eng = defaultdict(float)
        for row in self.res["deployed"]:
            if row["collection"] == "WT":
                eng[(row["line_key"], row["risk"])] += row["officers"]
        for row in self.res["packages"]:
            if row["collection"] == "WT":
                eng[(row["line_key"], row["risk"])] += row["units"]
        wbv = {(ws.cell(r, 1).value, risks[i]): ws.cell(r, 4 + i).value or 0
               for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value for i in range(4)}
        diffs = [(k, wbv.get(k, 0), eng.get(k, 0)) for k in set(wbv) | set(eng) if abs(wbv.get(k, 0) - eng.get(k, 0)) > 1e-9]
        self.assertFalse(diffs, diffs[:5])

    def test_summary_totals(self):
        ws = self.wb["Summary"]
        vals = {ws.cell(r, 2).value: ws.cell(r, 3).value for r in range(4, 12)}
        cal_cost = sum(self.engine_calendar("cost").values())
        self.assertAlmostEqual(vals["Match Day cost (GBP)"], cal_cost, places=2)
        self.assertAlmostEqual(vals["Wider Tournament cost (GBP)"], sum(self.res["wt_tournament"].values()), places=2)
        self.assertEqual(vals["Fixtures entered on the Demand Sheet"], len(build_sample.SYNTHETIC_FIXTURES))


if __name__ == "__main__":
    unittest.main()
