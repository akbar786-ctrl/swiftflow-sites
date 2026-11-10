"""
Builds EURO_2028_Power_BI_Reference.xlsx - the "master" lookup tables Power BI uses to turn
the forces' two collection sheets (Match Day Resource, Wider Tournament Resource) into one
tidy table and cost it exactly as the original collection workbook did.

  tblRowMapMD / tblRowMapWT   what every Excel row of the two sheets is: section, resource,
                              supplied by, line type, rank, PSU link, day rule, rate to use
  tblColMapMD / tblColMapWT   what every numeric column is: risk level, day, measure
  tblRates                    hourly rate by Supplied by + Rank (June 2028, original workbook)
  tblSettings                 mutual aid PSU / Serial package per unit per day
  tblWTDays                   Wider Tournament days at each risk level
  tblFixtures                 fixture list (fixture, date, risk level, force)

Every rule here is the original workbook's (see euro2028/v1_lift_shift/build_calc_pages.py,
which reproduced the original to 0 differences).

Usage: python3 make_reference.py [out.xlsx] [--fixtures csv-like "code,yyyy-mm-dd,risk,force;..."]
"""
import datetime as dt
import os
import re
import sys
import warnings

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
V1 = os.path.join(HERE, "..", "v1_lift_shift")
sys.path.insert(0, V1)
import build_calc_pages as B  # noqa: E402
import layout_v1 as LY  # noqa: E402

SOURCE = os.path.join(V1, "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "output", "EURO_2028_Power_BI_Reference.xlsx")
FS, MA = "Force supplied", "Mutual aid"
RISKS = ["Low", "Medium", "High", "Super High"]
DAYS = ["MD-2", "MD-1", "MD", "MD+1"]
MEASURES = ["Number", "Normal ToD", "Extra Hours"]

# June 2028 rates (Baseline Costs column X of the original workbook)
MA_RATE = {"Chief Superintendent": 106.654868085637, "Superintendent": 92.3808239153443,
           "Chief Inspector": 74.9030742125957, "Inspector": 67.3985037187601, "Sergeant": 67.4643332844955,
           "Constable": 54.2216189773909}
FS_RATE = {"Sergeant": 57.081492, "Constable": 48.233484}
PSU_PACKAGE = 35716.1613354
WT_DAYS = {"Low": 7, "Medium": 11, "High": 11, "Super High": 3}
RANKS = ["ACC", "Chief Superintendent", "Superintendent", "Chief Inspector", "Inspector", "Sergeant", "Constable",
         "Detective Inspector", "Detective Sergeant", "Detective Constable"]
DET = {"Detective Inspector": "Inspector", "Detective Sergeant": "Sergeant", "Detective Constable": "Constable"}
CELL_RATE = {"$X$10": (MA, "Chief Superintendent"), "$X$11": (MA, "Superintendent"), "$X$12": (MA, "Chief Inspector"),
             "$X$13": (MA, "Inspector"), "$X$14": (MA, "Sergeant"), "$X$15": (MA, "Constable"),
             "$X$31": (FS, "Sergeant"), "$X$32": (FS, "Constable")}


def clean_resource(label):
    s = (label or "").replace("\xa0", " ").strip()
    s = re.sub(r"^(Force Supplied|Mutual Aid Requested|Mutual Aid)\s*-?\s*", "", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip(" -")


def supply_of(label):
    s = (label or "").strip().lower()
    return MA if s.startswith("mutual aid") else FS if s.startswith("force supplied") else ""


def rate_key(t, sheet, r, in_stadium=False):
    """(rate supplied by, rate rank or '' = rank from column C) - the original workbook's rate for the row."""
    expr = B.rate_expr(t, sheet, r, in_stadium)
    if expr is None:
        return supply_of(t.label(sheet, r)), "NONE"
    if "VLOOKUP" in expr:
        return supply_of(t.label(sheet, r)), ""            # rank chosen in column C
    return CELL_RATE[re.search(r"\$X\$\d+", expr).group(0)]


def section_for(t, sheet, sections, r):
    prev = [s for s in sections if s < r]
    if not prev:
        return ""
    title = t.section_title(sheet, max(prev))
    return re.sub(r"\s+", " ", title or "").strip()


def row_map(t, sheet):
    md = sheet == LY.MD_SHEET
    config = LY.MD_CONFIG if md else LY.WT_CONFIG
    counts = LY.MD_COUNTS if md else LY.WT_COUNTS
    officers = LY.md_officer_rows() if md else LY.wt_officer_rows()
    roles = LY.MD_ROLE_ROWS if md else LY.WT_ROLE_ROWS
    rank_dd = LY.MD_RANK_DROPDOWN_ROWS if md else set()
    adds = LY.MD_ADDITIONAL_ROWS if md else LY.WT_ADDITIONAL_ROWS
    sections = LY.MD_SECTIONS if md else LY.WT_SECTIONS
    rows = []
    for c in sorted(config):
        count_row = next(k for k, v in counts.items() if v[-1] == c)
        group = counts[count_row][1] if md else "ALL"
        for i, rank in enumerate(["Inspector", "Sergeant", "Constable"]):
            r = c + i
            in_st = md and c in LY.MD_IN_STADIUM_CONFIG
            rs, rr = rate_key(t, sheet, r, in_st)
            rows.append(dict(Row=r, Section="PSU / Serial configuration", Resource=clean_resource(t.label(sheet, count_row)),
                             Supply=FS, LineType="PSU config", Rank=rank, RankFromC="No", RoleFromC="No",
                             Group=c, DayRule=group, RateSupply=rs if rs else FS,
                             RateRank=rr if rr != "NONE" else rank))
    for r in sorted(counts):
        sup = counts[r][0]
        group = counts[r][1] if md else "ALL"
        rows.append(dict(Row=r, Section=section_for(t, sheet, sections, r), Resource=clean_resource(t.label(sheet, r)),
                         Supply=FS if sup == "FS" else MA, LineType="PSU count" if sup == "FS" else "MA PSU count",
                         Rank="PSUs / Serials", RankFromC="No", RoleFromC="Yes", Group=counts[r][-1] or "",
                         DayRule=group, RateSupply="", RateRank=""))
    for r in officers:
        label = t.label(sheet, r)
        rs, rr = rate_key(t, sheet, r)
        if r in adds:
            kind, rank_c, role_c = "Additional", "No", "Yes"
        elif r in roles:
            kind, rank_c, role_c = "Role", "No", "Yes"
        elif r in rank_dd:
            kind, rank_c, role_c = "Officer", "Yes", "No"
        else:
            kind, rank_c, role_c = "Officer", "No", "No"
        rank = "" if rank_c == "Yes" or kind != "Officer" else t.rank(sheet, r)
        rows.append(dict(Row=r, Section=section_for(t, sheet, sections, r),
                         Resource=clean_resource(label) or "Additional Resources", Supply=supply_of(label),
                         LineType=kind, Rank=rank, RankFromC=rank_c, RoleFromC=role_c, Group="",
                         DayRule="STD" if md else "ALL", RateSupply=rs, RateRank="" if rr == "NONE" else rr))
        if rr == "NONE":
            rows[-1]["RateRank"] = rank                    # e.g. force supplied Inspector -> rate 0 in tblRates
    return sorted(rows, key=lambda x: x["Row"])


ROWMAP_HEAD = ["Excel row", "Section", "Resource", "Supplied by", "Line type", "Rank", "Rank from column C",
               "Role from column C", "PSU group", "Day rule", "Rate supplied by", "Rate rank"]
ROWMAP_KEYS = ["Row", "Section", "Resource", "Supply", "LineType", "Rank", "RankFromC", "RoleFromC", "Group",
               "DayRule", "RateSupply", "RateRank"]


def col_maps():
    md = [(4 + s * 12 + d * 3 + m, L(4 + s * 12 + d * 3 + m), RISKS[s], DAYS[d], MEASURES[m])
          for s in range(4) for d in range(4) for m in range(3)]
    wt = [(4 + s * 3 + m, L(4 + s * 3 + m), RISKS[s], "WT", MEASURES[m]) for s in range(4) for m in range(3)]
    return sorted(md), sorted(wt)


def rates():
    out = []
    for rank in RANKS:
        base = DET.get(rank, rank)
        out.append((FS, rank, FS_RATE.get(base, 0.0),
                    "Force supplied overtime rate" if base in FS_RATE else
                    "Force supplied Inspector and above: not costed (original workbook)"))
        if base in MA_RATE:
            out.append((MA, rank, MA_RATE[base], "Mutual aid hourly rate"))
    return out


def add_table(ws, name, top, headers, rows, widths, title, money=(), inputs=()):
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=12, color="1F3864")
    ws["A2"] = "OFFICIAL - Power BI reads the table below by name. Keep the headings; edit yellow cells only."
    ws["A2"].font = Font(italic=True, size=9)
    for j, h in enumerate(headers, 1):
        c = ws.cell(top, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        ws.column_dimensions[L(j)].width = widths[j - 1]
    rows = rows or [tuple(None for _ in headers)]
    for i, rw in enumerate(rows, 1):
        for j, v in enumerate(rw, 1):
            c = ws.cell(top + i, j, v)
            if j in money:
                c.number_format = "#,##0.00"
            if j in inputs:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
            if isinstance(v, dt.date):
                c.number_format = "dd/mm/yyyy"
    t = Table(displayName=name, ref=f"A{top}:{L(len(headers))}{top + len(rows)}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)


def build(out, fixtures=()):
    t = LY.Template(SOURCE)
    assert not LY.verify(SOURCE)
    wb = Workbook()
    first = True
    for name, sheet, title in (("tblRowMapMD", LY.MD_SHEET, "Row map - Match Day Resource"),
                               ("tblRowMapWT", LY.WT_SHEET, "Row map - Wider Tournament Resource")):
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = "RowMap MD" if sheet == LY.MD_SHEET else "RowMap WT"
        rows = [tuple(r[k] for k in ROWMAP_KEYS) for r in row_map(t, sheet)]
        add_table(ws, name, 4, ROWMAP_HEAD, rows, [10, 34, 52, 15, 14, 20, 12, 12, 10, 9, 15, 20],
                  title + " (what each Excel row of the forces' sheet is - do not edit unless the template changes)")
    md_cols, wt_cols = col_maps()
    ws = wb.create_sheet("ColMap MD")
    add_table(ws, "tblColMapMD", 4, ["Column number", "Column", "Risk", "Day", "Measure"], md_cols,
              [14, 10, 12, 8, 14], "Column map - Match Day Resource (D:AY)")
    ws = wb.create_sheet("ColMap WT")
    add_table(ws, "tblColMapWT", 4, ["Column number", "Column", "Risk", "Day", "Measure"], wt_cols,
              [14, 10, 12, 8, 14], "Column map - Wider Tournament Resource (D:O)")
    ws = wb.create_sheet("Rates")
    add_table(ws, "tblRates", 4, ["Supplied by", "Rank", "Hourly rate", "Basis"], rates(), [16, 24, 14, 64],
              "Hourly rates, June 2028 (original workbook, Baseline Costs column X)", money=(3,), inputs=(3,))
    ws = wb.create_sheet("Settings")
    add_table(ws, "tblSettings", 4, ["Key", "Setting", "Value"],
              [("MA_PSU_PACKAGE", "Mutual aid PSU / Serial package - cost per unit per day", PSU_PACKAGE)],
              [18, 60, 16], "Single values", money=(3,), inputs=(3,))
    ws = wb.create_sheet("WT days")
    add_table(ws, "tblWTDays", 4, ["Risk", "Days"], list(WT_DAYS.items()), [14, 10],
              "Wider Tournament - days at each risk level (national assumption)", inputs=(2,))
    ws = wb.create_sheet("Fixtures")
    add_table(ws, "tblFixtures", 4, ["Fixture", "Date", "Risk", "Force"], list(fixtures), [12, 14, 14, 34],
              "Fixtures - one line per fixture per force. Force exactly as typed on the forms (or the file name).",
              inputs=(1, 2, 3, 4))
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    args = sys.argv[1:]
    fx = []
    if "--fixtures" in args:
        i = args.index("--fixtures")
        for part in args[i + 1].split(";"):
            code, d, risk, force = part.split(",")
            fx.append((code, dt.date.fromisoformat(d), risk, force))
        args = args[:i] + args[i + 2:]
    print(build(os.path.abspath(args[0] if args else OUT), fx))
