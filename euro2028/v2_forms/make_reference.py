"""
Builds the reference-data workbook Power BI reads alongside the forces' forms.
Forces never see it. Edit the yellow cells when rates or assumptions change.

  tblRates          hourly rate by Supplied by + Rank (June 2028)
  tblRateOverrides  exceptions for a specific resource (e.g. In Stadium PSU Inspectors)
  tblSettings       single values (mutual aid PSU / Serial package)
  tblWTDays         number of Wider Tournament days at each risk level
  tblFixtures       fixture list: one line per fixture per force involved

Rates are the June 2028 figures (column X) of 'Baseline Costs' in the original collection
workbook; the rules (force-supplied Inspector and above not costed, In Stadium PSU
Inspectors at the mutual aid Inspector rate, mutual aid PSU package per unit per day,
WT days 7/11/11/3) are the original workbook's.

Usage: python3 make_reference.py [out.xlsx] [--sample-fixtures]
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import spec as S  # noqa: E402

OUT = os.path.join(HERE, "output", "EURO_2028_Reference_Data.xlsx")

MA_RATE = {"Chief Superintendent": 106.654868085637, "Superintendent": 92.3808239153443,
           "Chief Inspector": 74.9030742125957, "Inspector": 67.3985037187601, "Sergeant": 67.4643332844955,
           "Constable": 54.2216189773909}
FS_RATE = {"Sergeant": 57.081492, "Constable": 48.233484}
PSU_PACKAGE = 35716.1613354
WT_DAYS = {"Low": 7, "Medium": 11, "High": 11, "Super High": 3}
DETECTIVE = {"Detective Inspector": "Inspector", "Detective Sergeant": "Sergeant", "Detective Constable": "Constable"}
ALL_RANKS = ["Assistant Chief Constable"] + S.RANKS + list(DETECTIVE)


def rates():
    out = []
    for rank in ALL_RANKS:
        base = DETECTIVE.get(rank, rank)
        fs = FS_RATE.get(base, 0.0)
        out.append((S.FS, rank, fs, "Force supplied: Sergeant / Constable overtime rate" if base in FS_RATE
                    else "Force supplied Inspector and above: not costed (original workbook rule)"))
        if base in MA_RATE:
            out.append((S.MA, rank, MA_RATE[base], "Mutual aid hourly rate"))
    return out


OVERRIDES = [(f"In Stadium Public Order Policing (Matchday Only) - {lvl}", S.FS, "Inspector", MA_RATE["Inspector"],
              "Original workbook charged In Stadium PSU Inspectors at the mutual aid Inspector rate")
             for lvl in ("Level 2", "Level 3")]

SAMPLE_FIXTURES = [("A1", dt.date(2028, 6, 9), "High", "Sample Force A"),
                   ("A2", dt.date(2028, 6, 12), "Medium", "Sample Force A"),
                   ("A3", dt.date(2028, 6, 15), "Low", "Sample Force A"),
                   ("B1", dt.date(2028, 6, 10), "Medium", "Sample Force B"),
                   ("B2", dt.date(2028, 6, 11), "Super High", "Sample Force B"),
                   ("C1", dt.date(2028, 6, 20), "High", "Sample Force C"),
                   ("C1", dt.date(2028, 6, 20), "High", "Sample Force B")]


def table(ws, name, top, headers, rows, widths, money_cols=(), input_cols=()):
    for j, h in enumerate(headers, 1):
        c = ws.cell(top, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        ws.column_dimensions[c.column_letter].width = widths[j - 1]
    rows = rows or [tuple(None for _ in headers)]
    for i, rw in enumerate(rows, 1):
        for j, v in enumerate(rw, 1):
            c = ws.cell(top + i, j, v)
            if j in money_cols:
                c.number_format = "#,##0.00"
            if j in input_cols:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
            if isinstance(v, dt.date):
                c.number_format = "dd/mm/yyyy"
    t = Table(displayName=name, ref=f"A{top}:{ws.cell(top + len(rows), len(headers)).coordinate}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)


def note(ws, text):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=12, color="1F3864")
    ws["A2"] = "OFFICIAL - Power BI reads the table below. Change the yellow cells only; keep the column headings."
    ws["A2"].font = Font(italic=True, size=9)
    ws["A2"].alignment = Alignment(wrap_text=False)


def build(out, sample_fixtures=False):
    wb = Workbook()
    ws = wb.active
    ws.title = "Rates"
    note(ws, "Hourly rates by rank, June 2028 (from Baseline Costs column X of the original workbook)")
    table(ws, "tblRates", 4, ["Supplied by", "Rank", "Hourly rate", "Basis"], rates(), [16, 28, 14, 70],
          money_cols=(3,), input_cols=(3,))
    ws = wb.create_sheet("Rate overrides")
    note(ws, "Exceptions: a different hourly rate for one resource / supplied by / rank")
    table(ws, "tblRateOverrides", 4, ["Resource", "Supplied by", "Rank", "Hourly rate", "Reason"], OVERRIDES,
          [58, 16, 14, 14, 80], money_cols=(4,), input_cols=(4,))
    ws = wb.create_sheet("Settings")
    note(ws, "Single values used in the calculations")
    table(ws, "tblSettings", 4, ["Key", "Setting", "Value"],
          [("MA_PSU_PACKAGE", "Mutual aid PSU / Serial package - cost per unit per day (original workbook)",
            PSU_PACKAGE)], [18, 80, 16], money_cols=(3,), input_cols=(3,))
    ws = wb.create_sheet("WT days")
    note(ws, "Wider Tournament: number of days at each risk level (national assumption)")
    table(ws, "tblWTDays", 4, ["Risk", "Days"], list(WT_DAYS.items()), [16, 10], input_cols=(2,))
    ws = wb.create_sheet("Fixtures")
    note(ws, "Fixtures: one line per fixture per force involved. Risk = Low / Medium / High / Super High. "
             "Force must be written exactly as on the forms.")
    table(ws, "tblFixtures", 4, ["Fixture", "Date", "Risk", "Force"], SAMPLE_FIXTURES if sample_fixtures else [],
          [12, 14, 14, 34], input_cols=(1, 2, 3, 4))
    wb.save(out)
    return out


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = os.path.abspath(args[0] if args else OUT)
    print(build(out, "--sample-fixtures" in sys.argv))
