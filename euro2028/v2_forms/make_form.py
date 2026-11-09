"""
Builds the two-page EURO 2028 collection form that forces fill in:

  MD Collection   force details + what is needed for ONE fixture at each risk level,
                  Match Day -2 / -1 / Match Day / +1, Officers / Normal hrs / Extra hrs
  WT Collection   what is needed on a typical Wider Tournament day at each risk level

Each page is a single Excel table (tblMatchDay, tblWiderTournament; force details in
tblDetails) with one header row, so Power BI reads it with no clean-up. Drop-downs are
held in the cells themselves (no hidden list sheet). Cells that do not apply to a line
are grey and refuse entries.

Usage: python3 make_form.py [out.xlsx] [--sample NAME SEED]   (--sample fills synthetic test data)
"""
import os
import random
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import spec as S  # noqa: E402

OUT = os.path.join(HERE, "output", "EURO_2028_Resource_Collection_Form.xlsx")
FONT = "Arial"
NAVY = "1F3864"
GREY_LABEL, GREY_BLOCK, YELLOW, WHITE = "F2F2F2", "BFBFBF", "FFF2CC", "FFFFFF"
RISK_COL = {"Low": ("548235", "C6E0B4", "E2EFDA"), "Medium": ("BF8F00", "FFE699", "FFF2CC"),
            "High": ("C00000", "F8CBAD", "FCE4D6"), "Super High": ("7030A0", "D9C3E9", "EDE2F6")}
thin = Side(style="thin", color="A6A6A6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
KEY_W = {"Section": 26, "Resource": 46, "Supplied by": 15, "Role / Skill": 34, "Rank": 24}
FIRST = 6                            # first numeric column (F)


def fill(c):
    return PatternFill("solid", fgColor=c)


def style(cell, bg=None, bold=False, color="000000", size=9, wrap=False, h=None, border=True):
    cell.font = Font(name=FONT, size=size, bold=bold, color=color)
    if bg:
        cell.fill = fill(bg)
    cell.alignment = Alignment(wrap_text=wrap, horizontal=h, vertical="center")
    if border:
        cell.border = BOX


class DVs:
    """Collects cells per validation rule and adds one DataValidation per rule."""

    def __init__(self, ws):
        self.ws, self.rules = ws, {}

    def add(self, key, coord, **kw):
        if key not in self.rules:
            self.rules[key] = DataValidation(allow_blank=True, showErrorMessage=True, showInputMessage=True, **kw)
        self.rules[key].add(coord)

    def commit(self):
        for dv in self.rules.values():
            self.ws.add_data_validation(dv)


def inline(items):
    text = ",".join(items)
    assert len(text) <= 255 and "," not in "".join(i.replace(",", "") for i in items if "," in i)
    return f'"{text}"'


def number_dv(dvs, coord, measure):
    if measure == "Officers":
        dvs.add("officers", coord, type="whole", operator="between", formula1="0", formula2="9999",
                errorTitle="Whole number", error="Enter a whole number of officers (or units), 0 to 9999.",
                promptTitle="Officers", prompt="Number of officers needed. On 'Number of PSUs / Serials' lines: number of units.")
    else:
        dvs.add("hours", coord, type="decimal", operator="between", formula1="0", formula2="24",
                errorTitle="Hours", error="Enter hours per officer for the day, 0 to 24.",
                promptTitle=measure, prompt=("Normal tour of duty hours per officer." if measure == "Normal hrs"
                                             else "Extra (overtime) hours per officer."))


def blocked_dv(dvs, coord, why):
    dvs.add("blocked:" + why, coord, type="custom", formula1="FALSE", errorTitle="Not needed", error=why)


def allowed(r, day, measure):
    """Is this numeric cell used for this line?"""
    if r["rank"] == S.UNITS and measure != "Officers":
        return False, "Hours are not needed on a 'Number of PSUs / Serials' line - give hours on the officer lines."
    if day == "MD" and r["days"] == "NON_MD":
        return False, "This line is for days other than Match Day."
    if day not in (None, "MD") and r["days"] == "MD_ONLY":
        return False, "This line is for Match Day only."
    return True, ""


def header_block(ws, first_row, measures_cols, with_days):
    """Risk banner (and day banner) above the table header."""
    col = FIRST
    for risk in S.RISKS:
        dark, mid, light = RISK_COL[risk]
        span = len(S.DAYS) * 3 if with_days else 3
        ws.merge_cells(start_row=first_row, start_column=col, end_row=first_row, end_column=col + span - 1)
        c = ws.cell(first_row, col, risk.upper() + " RISK")
        style(c, bg=dark, bold=True, color="FFFFFF", size=11, h="center")
        for i in range(span):
            ws.cell(first_row, col + i).fill = fill(dark)
            ws.cell(first_row, col + i).border = BOX
        if with_days:
            for d, day in enumerate(S.DAYS):
                c0 = col + d * 3
                ws.merge_cells(start_row=first_row + 1, start_column=c0, end_row=first_row + 1, end_column=c0 + 2)
                c = ws.cell(first_row + 1, c0, S.DAY_NAMES[day])
                style(c, bg=mid, bold=True, size=9, h="center")
                for i in range(3):
                    ws.cell(first_row + 1, c0 + i).fill = fill(mid)
                    ws.cell(first_row + 1, c0 + i).border = BOX
        col += span


def page(ws, rows, headers, with_days, top, table_name, title, instructions):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws["A1"] = title
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color=NAVY)
    ws["F1"] = "OFFICIAL"
    ws["F1"].font = Font(name=FONT, size=11, bold=True, color="C00000")
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=5)
    ws["A4"] = instructions
    style(ws["A4"], size=9, wrap=True, border=False)
    ws.row_dimensions[4].height = 92
    key = [("Type a number", WHITE), ("Choose from the list", YELLOW), ("Not needed (locked)", GREY_BLOCK)]
    for i, (txt, bg) in enumerate(key):
        c = ws.cell(4, FIRST + i * 3, txt)
        ws.merge_cells(start_row=4, start_column=FIRST + i * 3, end_row=4, end_column=FIRST + i * 3 + 2)
        style(c, bg=bg, bold=True, size=9, h="center", wrap=True)
    hdr_row = top + (2 if with_days else 1)
    header_block(ws, top, len(headers), with_days)
    if with_days:
        ws.merge_cells(start_row=top, start_column=1, end_row=top + 1, end_column=5)
    else:
        ws.merge_cells(start_row=top, start_column=1, end_row=top, end_column=5)
    style(ws.cell(top, 1, "WHAT IS NEEDED"), bg=NAVY, bold=True, color="FFFFFF", size=11, h="center")
    cols = S.KEY_COLS + headers + ["Notes"]
    for j, name in enumerate(cols, 1):
        c = ws.cell(hdr_row, j, name)
        if j < FIRST:
            style(c, bg=NAVY, bold=True, color="FFFFFF", size=10, h="center", wrap=True)
            ws.column_dimensions[L(j)].width = KEY_W[name]
        elif name == "Notes":
            style(c, bg=NAVY, bold=True, color="FFFFFF", size=10, h="center")
            ws.column_dimensions[L(j)].width = 40
        else:
            risk = name.split(" | ")[0]
            style(c, bg=RISK_COL[risk][2], bold=True, size=8, h="center", wrap=True)
            ws.column_dimensions[L(j)].width = 9.5
    ws.row_dimensions[hdr_row].height = 48
    dvs = DVs(ws)
    r0 = hdr_row + 1
    for i, rw in enumerate(rows):
        r = r0 + i
        vals = [rw["section"], rw["resource"], rw["supply"] or "", rw["role"] if rw["role"] == S.NA else "",
                "" if rw["rank"] == S.RANK_DD else rw["rank"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v if v != "" else None)
            style(c, bg=GREY_LABEL, size=9, color="000000" if j > 2 else "404040")
        coord = lambda j: f"{L(j)}{r}"  # noqa: E731
        if rw["free"]:
            style(ws.cell(r, 2), bg=WHITE, size=9)
            dvs.add("free_resource", coord(2), type="textLength", operator="lessThanOrEqual", formula1="120",
                    promptTitle="Resource", prompt="Describe the additional resource (e.g. Communications Officer).")
        if rw["resource_list"]:
            style(ws.cell(r, 2), bg=YELLOW, size=9)
            dvs.add("res:" + rw["resource_list"], coord(2), type="list",
                    formula1=inline(S.RESOURCE_LISTS[rw["resource_list"]]), promptTitle="Resource",
                    prompt="Choose where this resource is deployed.")
        if rw["supply"] is None:
            style(ws.cell(r, 3), bg=YELLOW, size=9)
            dvs.add("supply", coord(3), type="list", formula1=inline([S.FS, S.MA]), promptTitle="Supplied by",
                    prompt="Force supplied or Mutual aid?")
        if rw["role"] != S.NA:
            style(ws.cell(r, 4), bg=YELLOW, size=9)
            dvs.add("role:" + rw["role"], coord(4), type="list", formula1=inline(S.ROLE_LISTS[rw["role"]]),
                    promptTitle="Role / Skill", prompt=("Choose the unit type." if rw["role"].startswith("PSU")
                                                        else "Choose the role or skill."))
        if rw["rank"] == S.RANK_DD:
            style(ws.cell(r, 5), bg=YELLOW, size=9)
            dvs.add("rank", coord(5), type="list", formula1=inline(S.RANKS), promptTitle="Rank",
                    prompt="Choose the rank. Use one line per rank.")
        for j, name in enumerate(headers, FIRST):
            parts = name.split(" | ")
            day, measure = (parts[1], parts[2]) if with_days else (None, parts[1])
            ok, why = allowed(rw, day, measure)
            c = ws.cell(r, j)
            if ok:
                style(c, bg=WHITE, size=9, h="center")
                c.number_format = "0" if measure == "Officers" else "0.0"
                number_dv(dvs, coord(j), measure)
            else:
                style(c, bg=GREY_BLOCK, size=9)
                blocked_dv(dvs, coord(j), why)
        style(ws.cell(r, FIRST + len(headers)), bg=WHITE, size=9, wrap=True)
    last_row, last_col = r0 + len(rows) - 1, FIRST + len(headers)
    dvs.commit()
    t = Table(displayName=table_name, ref=f"A{hdr_row}:{L(last_col)}{last_row}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False, showColumnStripes=False)
    ws.add_table(t)
    # repeated Section / Resource text shown faint so each group reads as a block
    for col in ("A", "B"):
        ws.conditional_formatting.add(
            f"{col}{r0 + 1}:{col}{last_row}",
            FormulaRule(formula=[f'AND(${col}{r0 + 1}<>"",${col}{r0 + 1}=${col}{r0},$A{r0 + 1}=$A{r0})'],
                        font=Font(color="BFBFBF")))
    # thicker line where a new section starts
    ws.conditional_formatting.add(
        f"A{r0 + 1}:{L(last_col)}{last_row}",
        FormulaRule(formula=[f"$A{r0 + 1}<>$A{r0}"],
                    border=Border(top=Side(style="medium", color=NAVY))))
    ws.freeze_panes = f"{L(FIRST)}{r0}"
    ws.oddFooter.center.text = "OFFICIAL"
    ws.oddHeader.center.text = "OFFICIAL"
    return r0, last_row


MD_INSTRUCTIONS = (
    "Tell us what your force needs for ONE fixture at each risk level, for each day from Match Day -2 to "
    "Match Day +1. Officers = number of officers. Normal hrs = normal tour of duty hours per officer. "
    "Extra hrs = extra (overtime) hours per officer.\n"
    "PSUs / Serials: on the 'Number of PSUs / Serials' line give how many units and choose the unit type; on "
    "the Inspector / Sergeant / Constable lines give the TOTAL officers in those units.\n"
    "Role / Skill lines: choose the role and the rank (one line per rank). Leave anything you do not need blank."
)
WT_INSTRUCTIONS = (
    "Tell us what your force needs on a TYPICAL Wider Tournament day (no fixture in your area) at each risk "
    "level. The number of days at each risk level is set nationally.\n"
    "Officers / Normal hrs / Extra hrs and PSU / Serial lines work as on the MD Collection page.\n"
    "Your force details are entered once, at the top of the MD Collection page."
)


def build(out, sample=None):
    wb = Workbook()
    md = wb.active
    md.title = "MD Collection"
    # force details (one-row table, read by Power BI)
    det = ["Force", "Completed by", "Email", "Phone", "Date completed"]
    for j, h in enumerate(det, 1):
        style(md.cell(2, j, h), bg=NAVY, bold=True, color="FFFFFF", size=10, h="center")
        style(md.cell(3, j), bg=YELLOW if j == 1 else WHITE, size=10)
    md["E3"].number_format = "dd/mm/yyyy"
    md.row_dimensions[3].height = 20
    td = Table(displayName="tblDetails", ref="A2:E3")
    td.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    md.add_table(td)
    dv = DataValidation(type="textLength", operator="between", formula1="2", formula2="80", allow_blank=False,
                        showInputMessage=True, showErrorMessage=True, promptTitle="Force",
                        prompt="Your force name, written the same way every time (e.g. West Midlands Police).",
                        errorTitle="Force", error="Please enter your force name.")
    dv.add("A3")
    md.add_data_validation(dv)
    dv2 = DataValidation(type="date", operator="greaterThan", formula1="45658", allow_blank=True,
                         showErrorMessage=True, error="Enter a date, e.g. 14/11/2026.")
    dv2.add("E3")
    md.add_data_validation(dv2)
    md_r0, md_last = page(md, S.MD_ROWS, S.md_headers(), True, 5, "tblMatchDay",
                          "EURO 2028 - Match Day resource collection", MD_INSTRUCTIONS)
    wt = wb.create_sheet("WT Collection")
    wt["A2"] = "Force details are entered on the MD Collection page."
    wt["A2"].font = Font(name=FONT, size=10, italic=True)
    wt_r0, wt_last = page(wt, S.WT_ROWS, S.wt_headers(), False, 5, "tblWiderTournament",
                          "EURO 2028 - Wider Tournament resource collection", WT_INSTRUCTIONS)
    for ws in (md, wt):
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    md.sheet_view.selection[0].activeCell = "A3"
    md.sheet_view.selection[0].sqref = "A3"
    if sample:
        fill_sample(wb, *sample)
    wb.save(out)
    return out


# ------------------------------------------------------------- synthetic test data
def fill_sample(wb, name, seed):
    """Fills plausible-looking SYNTHETIC numbers (not real force data) for testing the Power BI model."""
    rnd = random.Random(seed)
    md, wt = wb["MD Collection"], wb["WT Collection"]
    md["A3"], md["B3"], md["E3"] = name, "Test user", None
    for ws, rows, headers, with_days, r0 in ((md, S.MD_ROWS, S.md_headers(), True, 8),
                                             (wt, S.WT_ROWS, S.wt_headers(), False, 7)):
        for i, rw in enumerate(rows):
            if rnd.random() < 0.55:
                continue                                   # most lines left blank, as in real returns
            r = r0 + i
            if rw["free"]:
                ws.cell(r, 2, rnd.choice(["Communications Officer", "Welfare Officer", "Event Liaison"]))
                ws.cell(r, 3, rnd.choice([S.FS, S.MA]))
            if rw["resource_list"]:
                ws.cell(r, 2, rnd.choice(S.RESOURCE_LISTS[rw["resource_list"]]))
            if rw["role"] != S.NA:
                ws.cell(r, 4, rnd.choice(S.ROLE_LISTS[rw["role"]]))
            if rw["rank"] == S.RANK_DD:
                ws.cell(r, 5, rnd.choice(S.RANKS[2:]))
            for j, h in enumerate(headers, FIRST):
                parts = h.split(" | ")
                day, measure = (parts[1], parts[2]) if with_days else (None, parts[1])
                if not allowed(rw, day, measure)[0] or rnd.random() < 0.3:
                    continue
                risk = S.RISKS.index(parts[0])
                if measure == "Officers":
                    base = 2 if rw["rank"] == S.UNITS else (1 if rw["rank"] != "Constable" else 6)
                    ws.cell(r, j, base * (risk + 1) + rnd.randint(0, 3))
                elif measure == "Normal hrs":
                    ws.cell(r, j, rnd.choice([8, 10]))
                else:
                    ws.cell(r, j, rnd.choice([0, 2, 4, 4.5, 6]))


if __name__ == "__main__":
    args = sys.argv[1:]
    sample = None
    if "--sample" in args:
        i = args.index("--sample")
        sample = (args[i + 1], int(args[i + 2]))
        args = args[:i] + args[i + 3:]
    out = os.path.abspath(args[0] if args else OUT)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    print(build(out, sample))
