"""
Builds the EURO 2028 force collection template (MD + WT) from catalogue.py.

Layout (matches the forces' draft form):
  A  Cell Input Instructions   (Rank dropdown on role-profile rows)
  B  Resource                  (free-text description on Additional Resource rows)
  C  Rank (Drop Down)          (Unit dropdown on PSU count rows, Role Profile dropdown on role rows)
  D… inputs, risk level on top: Low | Medium | High | Super High,
      then (MD only) Match Day -2 | -1 | Match Day | +1, then Number / Normal ToD / Extra Hours
  then automatic totals, then a HIDDEN key column.

Machine-readable contract (used by Power Query and the Python engine):
  * Row 6 (hidden) holds column keys: "<Risk>|<Day>|<Measure>" over input
    columns, "COLA"/"COLB"/"COLC" over A-C, "ROWKEY" over the key column.
  * The key column holds "<ResourceID>|<Line>" on data rows and
    "SECTION|<coll>-<section>" on section header rows.
  * Cover sheet table tblSectionStatus holds section completion status.
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, os.path.dirname(__file__))
import catalogue as C  # noqa: E402
import reference_data as R  # noqa: E402

NAVY, BLUE, LIGHT = "1F3864", "2F5597", "D9E1F2"
MA_GREY, NA_GREY, WHITE = "EDEDED", "BFBFBF", "FFFFFF"
RISK_TINT = {"L": "E2EFDA", "M": "FFF2CC", "H": "FCE4D6", "SH": "F4CCCC"}

thin = Side(style="thin", color="8EA9DB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_TITLE = Font(name="Aptos", size=18, bold=True, color=WHITE)
F_HDR = Font(name="Aptos", size=10, bold=True, color=WHITE)
F_BODY = Font(name="Aptos", size=9)
F_BOLD = Font(name="Aptos", size=9, bold=True)
F_NOTE = Font(name="Aptos", size=9, italic=True, color="595959")


def FILL(c):
    return PatternFill("solid", fgColor=c)


NA_FILL = PatternFill("lightGray", fgColor=NA_GREY, bgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ACROSS = Alignment(horizontal="centerContinuous", vertical="center")
UNLOCK = Protection(locked=False)
LOCK = Protection(locked=True)

FIRST_INPUT_COL = 4      # D
KEY_ROW = 6
DATA_START_ROW = 7
OFFICER_TYPES = ("OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL")


def input_columns(collection):
    """Ordered list of (col_index, risk, day, measure)."""
    cols, c = [], FIRST_INPUT_COL
    days = C.MD_DAYS if collection == "MD" else [("WT", "Per day", 0, 1)]
    for risk, _, _ in C.RISK_LEVELS:
        for day, *_ in days:
            for meas, _ in C.MEASURES:
                cols.append((c, risk, day, meas))
                c += 1
    return cols


def layout(collection):
    cols = input_columns(collection)
    last_in = cols[-1][0]
    tot_start = last_in + 2
    tot_cols = {risk: (tot_start + i * 2, tot_start + i * 2 + 1) for i, (risk, _, _) in enumerate(C.RISK_LEVELS)}
    key_col = tot_start + 2 * len(C.RISK_LEVELS) + 1
    return cols, last_in, tot_start, tot_cols, key_col


def _lists_sheet(wb):
    ws = wb.create_sheet(C.LISTS_SHEET)
    lists = {code: [C.RANK_LABEL[x] for x in codes] for code, codes in C.RANK_LISTS.items()}
    lists.update(C.UNIT_LISTS)
    lists.update(C.ROLE_LISTS)
    lists["SectionStatus"] = C.SECTION_STATUS
    lists["Forces"] = [f[0] for f in R.FORCES]
    for ci, (name, vals) in enumerate(lists.items(), 1):
        ws.cell(1, ci, name).font = F_BOLD
        for ri, v in enumerate(vals, 2):
            ws.cell(ri, ci, v)
        ref = f"'{C.LISTS_SHEET}'!${L(ci)}$2:${L(ci)}${len(vals) + 1}"
        wb.defined_names[f"lst_{name}"] = DefinedName(f"lst_{name}", attr_text=ref)
        ws.column_dimensions[L(ci)].width = 34
    ws.protection.sheet = True


def _cover(wb):
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    for col, w in (("A", 3), ("B", 30), ("C", 46), ("D", 22), ("E", 50)):
        ws.column_dimensions[col].width = w
    ws["B2"] = "EURO 2028 Resource Forecast Collection"
    ws["B2"].font = Font(name="Aptos", size=20, bold=True, color=NAVY)
    ws["B3"] = "NPoCC EURO 2028 Police Planning  |  OFFICIAL"
    ws["B3"].font = F_NOTE
    fields = [
        ("cv_Force", "Force", "Choose your force from the list."),
        ("cv_ContactName", "Contact name", ""),
        ("cv_ContactRole", "Contact role", ""),
        ("cv_ContactEmail", "Contact email", ""),
        ("cv_ContactPhone", "Contact phone", ""),
        ("cv_SubmissionDate", "Date submitted", "dd/mm/yyyy"),
        ("cv_SubmissionNo", "Submission number", "1 for the first submission; add 1 for each resubmission."),
        ("cv_Comments", "Comments", "Any assumptions the planning team should know about."),
    ]
    dv_force = DataValidation(type="list", formula1="lst_Forces", showErrorMessage=True,
                              error="Choose a force from the list.")
    dv_date = DataValidation(type="date", operator="between", formula1="DATE(2026,1,1)",
                             formula2="DATE(2028,12,31)", showErrorMessage=True,
                             error="Enter a date between 2026 and 2028.")
    dv_int = DataValidation(type="whole", operator="between", formula1="1", formula2="99",
                            showErrorMessage=True, error="Enter a whole number 1-99.")
    r = 5
    for name, label, hint in fields:
        ws.cell(r, 2, label).font = F_BOLD
        c = ws.cell(r, 3)
        c.border, c.protection = BORDER, UNLOCK
        if hint:
            ws.cell(r, 4, hint).font = F_NOTE
        wb.defined_names[name] = DefinedName(name, attr_text=f"Cover!$C${r}")
        if name == "cv_Force":
            dv_force.add(c.coordinate)
        elif name == "cv_SubmissionDate":
            dv_date.add(c.coordinate)
            c.number_format = "dd/mm/yyyy"
        elif name == "cv_SubmissionNo":
            dv_int.add(c.coordinate)
        elif name == "cv_Comments":
            ws.row_dimensions[r].height = 48
            c.alignment = WRAP
        r += 1
    ws.cell(r, 2, "Template version").font = F_BOLD
    ws.cell(r, 3, C.TEMPLATE_VERSION)
    wb.defined_names["cv_TemplateVersion"] = DefinedName("cv_TemplateVersion", attr_text=f"Cover!$C${r}")

    # Section completion checklist (read by the model: blank = 0 only in Complete sections)
    r += 2
    ws.cell(r, 2, "Section completion").font = Font(name="Aptos", size=12, bold=True, color=NAVY)
    ws.cell(r + 1, 2, "Set each section to Complete (blank cells count as 0) or Not required. "
                      "Sections left Not started / In progress are reported as incomplete.").font = F_NOTE
    hdr = r + 2
    for ci, h in enumerate(["Collection", "Section", "Status", "SectionKey"], 2):
        ws.cell(hdr, ci, h)
    dv_status = DataValidation(type="list", formula1="lst_SectionStatus", showErrorMessage=True)
    rr = hdr + 1
    for (coll, code), title in C.SECTIONS.items():
        ws.cell(rr, 2, "Match Day" if coll == "MD" else "Wider Tournament")
        ws.cell(rr, 3, title)
        st = ws.cell(rr, 4, "Not started")
        st.protection = UNLOCK
        dv_status.add(st.coordinate)
        ws.cell(rr, 5, f"{coll}-{code}").font = F_NOTE
        rr += 1
    t = Table(displayName="tblSectionStatus", ref=f"B{hdr}:E{rr - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    for dv in (dv_force, dv_date, dv_int, dv_status):
        ws.add_data_validation(dv)
    ws.protection.sheet = True


GUIDANCE = [
    ("h", "How to complete this collection"),
    ("p", "1. Complete the Cover sheet: force, contact details, and the Section completion list."),
    ("p", "2. 'Match Day Resource': enter what you need for ONE fixture at each risk level, for the day before "
          "last (Match Day -2), the day before (Match Day -1), Match Day and the day after (Match Day +1)."),
    ("p", "3. 'Wider Tournament Resource': enter what you need on ONE typical wider-tournament day at each risk level."),
    ("p", "4. Fixtures and opponents are not yet known, so complete every risk level: Low, Medium, High and "
          "Super High. The planning team applies your figures to fixtures."),
    ("h", "What to enter"),
    ("p", "Number required - how many officers (or PSUs / serials, or shifts where column A says so)."),
    ("p", "Normal ToD - the normal tour-of-duty hours EACH officer works that day (e.g. 10)."),
    ("p", "Extra Hours - the hours EACH officer works beyond their normal tour of duty (e.g. 4)."),
    ("h", "Blank, zero and not applicable"),
    ("p", "Leave a cell BLANK if not yet decided. Enter 0 if none are needed. Grey hatched cells are NOT "
          "APPLICABLE and are locked. In a section marked Complete on the Cover sheet, blanks count as 0."),
    ("h", "PSU / serial rows"),
    ("p", "Configuration rows: enter how many of each rank make up ONE Level 2 PSU or ONE Level 3 BDU / serial, "
          "and their hours. Public Order rows: enter how many PSUs / serials you need. The model multiplies them."),
    ("p", "Worked example: a Level 2 PSU of 1 Inspector, 3 Sergeants and 18 Constables, each working 10 hours "
          "normal ToD + 4 extra hours, and 6 PSUs on Match Day (High) gives 6 Inspectors, 18 Sergeants and "
          "108 Constables, 1,848 hours in total."),
    ("p", "Non-matchday rows cover Match Day -2, -1 and +1 only; matchday-only rows cover Match Day only."),
    ("p", "Mutual aid PSUs are charged as a package; enter numbers only."),
    ("h", "Drop-down rows"),
    ("p", "Rank (Drop Down) rows: choose the rank in column C. Role-profile rows: choose the rank in column A and "
          "the role profile in column C. Additional Resources: describe the resource in column B and choose the "
          "rank in column C."),
    ("p", "Protest removal: use the Protest Removal rows; do not repeat the same officers as role profiles."),
    ("h", "Checks"),
    ("p", "Numbers: whole numbers 0-9,999. Hours: 0-24. Amber = looks incomplete (number with no hours, or hours "
          "with no number). Red = rank or role not chosen, or more than 24 hours."),
    ("p", "Do not insert, delete or move rows or columns: the planning model reads the sheet automatically."),
]


def _guidance(wb):
    ws = wb.create_sheet("Guidance")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    ws.cell(2, 2, "Guidance").font = Font(name="Aptos", size=18, bold=True, color=NAVY)
    r = 4
    for kind, text in GUIDANCE:
        c = ws.cell(r, 2, text)
        if kind == "h":
            c.font = Font(name="Aptos", size=12, bold=True, color=BLUE)
            r += 1
        else:
            c.font = Font(name="Aptos", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * (len(text) // 115 + 1)
        r += 1
    r += 1
    ws.cell(r, 2, "Legend").font = Font(name="Aptos", size=12, bold=True, color=BLUE)
    for label, fill in [("Input (force supplied)", FILL(WHITE)), ("Input (mutual aid)", FILL(MA_GREY)),
                        ("Not applicable (locked)", NA_FILL)]:
        r += 1
        ws.cell(r, 2, label).fill = fill
        ws.cell(r, 2).border = BORDER
    ws.protection.sheet = True


def _collection(wb, coll):
    ws = wb.create_sheet(C.SHEET_NAMES[coll])
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    cols, last_in, tot_start, tot_cols, key_col = layout(coll)
    last_vis = tot_start + 2 * len(C.RISK_LEVELS) - 1
    for c, w in {1: 30, 2: 46, 3: 22}.items():
        ws.column_dimensions[L(c)].width = w
    for c, *_ in cols:
        ws.column_dimensions[L(c)].width = 10.5
    ws.column_dimensions[L(last_in + 1)].width = 2
    for c in range(tot_start, last_vis + 1):
        ws.column_dimensions[L(c)].width = 11
    ws.column_dimensions[L(last_vis + 1)].width = 2
    ws.column_dimensions[L(key_col)].hidden = True

    # ---- header
    title = ("Match Day" if coll == "MD" else "Wider Tournament") + " Resource Forecast Collection"
    ws.row_dimensions[1].height = 36
    for c in range(1, last_vis + 1):
        ws.cell(1, c).fill = FILL(NAVY)
        ws.cell(2, c).fill = FILL(LIGHT)
    ws["A1"] = f"EURO 2028  |  {title}"
    ws["A1"].font = F_TITLE
    ws["A2"] = ('="Force: "&IF(cv_Force="","(complete the Cover sheet)",cv_Force)&"   |   Template v'
                + C.TEMPLATE_VERSION + '   |   OFFICIAL"')
    ws["A2"].font = F_BOLD
    for c, h in enumerate(["Cell Input Instructions", "Resource", "Rank (Drop Down)"], 1):
        for rr in (3, 4, 5):
            ws.cell(rr, c).fill = FILL(NAVY)
            ws.cell(rr, c).border = BORDER
        ws.cell(5, c, h).font = F_HDR
        ws.cell(5, c).alignment = CENTER
    ws.row_dimensions[5].height = 30
    risk_label = {r: lbl for r, lbl, _ in C.RISK_LEVELS}
    day_label = {d: lbl for d, lbl, *_ in C.MD_DAYS}
    day_label["WT"] = "Per day (typical tournament day)"
    meas_label = dict(C.MEASURES)
    prev_risk = prev_day = None
    for c, risk, day, meas in cols:
        top = ws.cell(3, c)
        top.fill, top.font, top.border, top.alignment = FILL(BLUE), F_HDR, BORDER, ACROSS
        if risk != prev_risk:
            top.value = f"{risk_label[risk].upper()} RISK"
            prev_day = None
        d = ws.cell(4, c)
        d.fill, d.font, d.border, d.alignment = FILL(RISK_TINT[risk]), F_BOLD, BORDER, ACROSS
        if day != prev_day:
            d.value = day_label[day]
        m = ws.cell(5, c, meas_label[meas])
        m.fill, m.font, m.alignment, m.border = FILL(RISK_TINT[risk]), F_BOLD, CENTER, BORDER
        ws.cell(KEY_ROW, c, f"{risk}|{day}|{meas}")
        prev_risk, prev_day = risk, day
    ws.cell(3, tot_start, "AUTOMATIC TOTALS (all days)" if coll == "MD" else "AUTOMATIC TOTALS (per day)")
    for risk, (cn, ch) in tot_cols.items():
        for c, lbl in ((cn, "Total number"), (ch, "Total hours")):
            ws.cell(3, c).fill = FILL(NAVY)
            ws.cell(3, c).font = F_HDR
            ws.cell(3, c).alignment = ACROSS
            ws.cell(4, c).fill = FILL(RISK_TINT[risk])
            ws.cell(4, c).font = F_BOLD
            ws.cell(4, c).alignment = ACROSS
            ws.cell(5, c, lbl).font = F_BOLD
            ws.cell(5, c).fill = FILL(RISK_TINT[risk])
            ws.cell(5, c).alignment = CENTER
            for rr in (3, 4, 5):
                ws.cell(rr, c).border = BORDER
        ws.cell(4, cn, f"{risk_label[risk]} totals")
        ws.cell(KEY_ROW, cn, f"TOTAL|{risk}|NUM")
        ws.cell(KEY_ROW, ch, f"TOTAL|{risk}|HRS")
    for c, k in ((1, "COLA"), (2, "COLB"), (3, "COLC"), (key_col, "ROWKEY")):
        ws.cell(KEY_ROW, c, k)
    ws.row_dimensions[KEY_ROW].hidden = True

    # ---- validations
    dv_num = DataValidation(type="whole", operator="between", formula1="0", formula2="9999", showErrorMessage=True,
                            errorTitle="Number required", error="Enter a whole number 0-9,999, or leave blank.")
    dv_cfg = DataValidation(type="whole", operator="between", formula1="0", formula2="60", showErrorMessage=True,
                            errorTitle="Officers per PSU / serial",
                            error="Enter a whole number 0-60 (officers in ONE PSU / serial).")
    dv_hrs = DataValidation(type="decimal", operator="between", formula1="0", formula2="24", showErrorMessage=True,
                            errorTitle="Hours", error="Enter hours between 0 and 24, or leave blank.")
    dv_lists = {}

    def dv_list(name):
        if name not in dv_lists:
            dv_lists[name] = DataValidation(type="list", formula1=f"lst_{name}", showErrorMessage=True,
                                            error="Choose a value from the list.")
        return dv_lists[name]

    officer_rows, rank_rows_a, rank_rows_c, role_rows_c, desc_rows_b = [], [], [], [], []
    r = DATA_START_ROW
    for (sc_coll, sc_code), sc_title in [(k, t) for k, t in C.SECTIONS.items() if k[0] == coll]:
        sec_res = [x for x in C.RESOURCES if x["collection"] == coll and x["section_code"] == sc_code]
        types = {x["row_type"] for x in sec_res}
        for c in range(1, last_vis + 1):
            ws.cell(r, c).fill = FILL(BLUE)
            ws.cell(r, c).border = BORDER
        ws.cell(r, key_col, f"SECTION|{coll}-{sc_code}")
        ws.cell(r, 2, sc_title).font = F_HDR
        ws.cell(r, 2).alignment = Alignment(horizontal="center", vertical="center")
        if "ROLE" in types:
            ws.cell(r, 1, "Rank (Drop Down)").font = F_HDR
            ws.cell(r, 3, "Role Profiles (Drop Down)").font = F_HDR
        elif "ADDITIONAL" in types:
            ws.cell(r, 3, "Rank (Drop Down)").font = F_HDR
        elif "RANK_SELECT" in types:
            ws.cell(r, 3, "Rank (Drop Down)").font = F_HDR
        for c in (1, 3):
            ws.cell(r, c).alignment = CENTER
        ws.row_dimensions[r].height = 18
        r += 1
        for res in sec_res:
            res_lines = list(C.lines(res))
            rt = res["row_type"]
            first = r
            base = FILL(MA_GREY) if res["supply"] == "MA" else FILL(WHITE)
            for ln in res_lines:
                ws.cell(r, key_col, C.row_key(res["resource_id"], ln["line"]))
                for c in range(1, last_vis + 1):
                    cell = ws.cell(r, c)
                    cell.border, cell.font = BORDER, F_BODY
                    if c not in (last_in + 1,):
                        cell.fill = base
                a, b, cc = ws.cell(r, 1), ws.cell(r, 2), ws.cell(r, 3)
                a.alignment, b.alignment, cc.alignment = WRAP, WRAP, CENTER
                # column A
                if rt == "ROLE":
                    a.protection, a.fill, a.alignment = UNLOCK, FILL(WHITE), CENTER
                    dv_list(res["rank_list"]).add(a.coordinate)
                    rank_rows_a.append(r)
                elif res["basis"] == "Shifts":
                    a.value = "Number of shifts"
                    a.alignment = CENTER
                elif rt == "ADDITIONAL":
                    a.value = "Force Supplied" if res["supply"] == "FS" else "Mutual Aid Requested"
                    a.alignment = CENTER
                elif ln["line"] == 1:
                    a.value = res["instructions"] or None
                # column B
                if rt == "ADDITIONAL":
                    b.protection, b.fill = UNLOCK, FILL(WHITE)
                    desc_rows_b.append(r)
                elif ln["line"] == 1:
                    b.value = res["name"]
                # column C
                if rt in ("PSU_COUNT", "MA_PSU"):
                    cc.value, cc.protection = ln["label"], UNLOCK
                    dv_list(ln["unit_list"]).add(cc.coordinate)
                elif rt == "ROLE":
                    cc.protection, cc.fill = UNLOCK, FILL(WHITE)
                    dv_list(res["role_list"]).add(cc.coordinate)
                    role_rows_c.append(r)
                elif ln["rank"]:
                    cc.value = ln["label"]
                else:
                    cc.protection, cc.fill = UNLOCK, FILL(WHITE)
                    dv_list(res["rank_list"]).add(cc.coordinate)
                    rank_rows_c.append(r)
                # inputs
                for c, risk, day, meas in cols:
                    cell = ws.cell(r, c)
                    if C.applicable(res, meas, day if coll == "MD" else None):
                        cell.protection, cell.alignment = UNLOCK, CENTER
                        if meas == "NUM":
                            (dv_cfg if rt == "PSU_CONFIG" else dv_num).add(cell.coordinate)
                            cell.number_format = "0"
                        else:
                            dv_hrs.add(cell.coordinate)
                            cell.number_format = "0.0"
                    else:
                        cell.fill, cell.protection = NA_FILL, LOCK
                # totals
                for risk, (cn, ch) in tot_cols.items():
                    if rt == "PSU_CONFIG":
                        ws.cell(r, cn).fill = NA_FILL
                        ws.cell(r, ch).fill = NA_FILL
                        continue
                    rc = [x for x in cols if x[1] == risk]
                    nums = [f"{L(c)}{r}" for c, _, _, m in rc if m == "NUM"]
                    ws.cell(r, cn, f'=IF(COUNT({",".join(nums)})=0,"",SUM({",".join(nums)}))')
                    ws.cell(r, cn).number_format = "#,##0"
                    ws.cell(r, cn).alignment = CENTER
                    if rt in OFFICER_TYPES:
                        terms = []
                        for d in dict.fromkeys(dd for _, _, dd, _ in rc):
                            n_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "NUM")
                            t_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "TOD")
                            x_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "XH")
                            terms.append(f"N({n_}{r})*(N({t_}{r})+N({x_}{r}))")
                        ws.cell(r, ch, f'=IF(COUNT({",".join(nums)})=0,"",{"+".join(terms)})')
                        ws.cell(r, ch).number_format = "#,##0.0"
                        ws.cell(r, ch).alignment = CENTER
                    else:
                        ws.cell(r, ch).fill = NA_FILL
                if rt in OFFICER_TYPES:
                    officer_rows.append(r)
                r += 1
            if len(res_lines) > 1:
                # A (instructions) and B (resource) span the resource's lines, except where
                # A or B hold a per-line value (rank on role rows, description / supply on
                # additional rows, "Number of shifts" on shift rows).
                if rt == "ADDITIONAL":
                    merge_cols = []
                elif rt == "ROLE" or res["basis"] == "Shifts":
                    merge_cols = [2]
                else:
                    merge_cols = [1, 2]
                for c in merge_cols:
                    ws.merge_cells(start_row=first, start_column=c, end_row=r - 1, end_column=c)
                    ws.cell(first, c).alignment = WRAP
        r += 1

    # ---- conditional formatting
    def rows_ref(col_a, col_b, rows):
        out, start, prev = [], None, None
        for x in sorted(rows):
            if start is None:
                start = prev = x
            elif x == prev + 1:
                prev = x
            else:
                out.append(f"{col_a}{start}:{col_b}{prev}")
                start = prev = x
        if start is not None:
            out.append(f"{col_a}{start}:{col_b}{prev}")
        return " ".join(out)

    amber, red = FILL("FFD966"), FILL("FF7C80")
    if officer_rows:
        f0 = min(officer_rows)
        for c, risk, day, meas in cols:
            if meas != "NUM":
                continue
            n_, t_, x_ = L(c), L(c + 1), L(c + 2)
            sq = rows_ref(n_, x_, officer_rows)
            ws.conditional_formatting.add(sq, FormulaRule(
                formula=[f'AND(N(${n_}{f0})>0,${t_}{f0}="",${x_}{f0}="")'], fill=amber))
            ws.conditional_formatting.add(sq, FormulaRule(
                formula=[f'AND(${n_}{f0}="",OR(${t_}{f0}<>"",${x_}{f0}<>""))'], fill=amber))
            ws.conditional_formatting.add(sq, FormulaRule(formula=[f"N(${t_}{f0})+N(${x_}{f0})>24"], fill=red))
    rng = f"${L(FIRST_INPUT_COL)}{{r}}:${L(last_in)}{{r}}"
    for col, rows in (("A", rank_rows_a), ("C", rank_rows_c), ("C", role_rows_c), ("B", desc_rows_b)):
        if rows:
            f0 = min(rows)
            ws.conditional_formatting.add(rows_ref(col, col, rows), FormulaRule(
                formula=[f'AND(${col}{f0}="",SUM({rng.format(r=f0)})>0)'], fill=red))
    for dv in [dv_num, dv_cfg, dv_hrs, *dv_lists.values()]:
        if dv.sqref:
            ws.add_data_validation(dv)

    ws.freeze_panes = ws.cell(DATA_START_ROW, FIRST_INPUT_COL)
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.print_title_rows = "3:5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def _optional_tables(wb):
    specs = [
        ("Command Team", "tblCommandTeam",
         ["Rank", "Number of posts", "Start month", "End month", "Notes"], 20,
         "Planning / command team posts and the months they are required (legacy 'Command Team & Planning'). "
         "End month is the LAST month the post is required (inclusive)."),
        ("Other Resources", "tblOtherResources",
         ["Category", "Rank", "Hours FY 2027-28", "Hours FY 2028-29", "Notes"], 20,
         "Other overtime hours by category (legacy 'Other Resources')."),
    ]
    for sheet, tname, cols, n, note in specs:
        ws = wb.create_sheet(sheet)
        ws.sheet_view.showGridLines = False
        ws["A1"] = sheet
        ws["A1"].font = Font(name="Aptos", size=16, bold=True, color=NAVY)
        ws["A2"] = note
        ws["A2"].font = F_NOTE
        for ci, h in enumerate(cols, 1):
            ws.cell(4, ci, h)
            ws.column_dimensions[L(ci)].width = 34 if ci == 1 else 22
        for rr in range(5, 5 + n):
            for ci in range(1, len(cols) + 1):
                ws.cell(rr, ci).protection = UNLOCK
        if sheet == "Other Resources":
            for i, cat in enumerate(R.OTHER_RESOURCE_CATEGORIES):
                for j, rk in enumerate(["Sergeant", "Constable"]):
                    ws.cell(5 + i * 2 + j, 1, cat)
                    ws.cell(5 + i * 2 + j, 2, rk)
        else:
            for i, rk in enumerate(["Chief Superintendent", "Superintendent", "Chief Inspector", "Inspector",
                                    "Sergeant", "Constable"]):
                ws.cell(5 + i, 1, rk)
        t = Table(displayName=tname, ref=f"A4:{L(len(cols))}{4 + n}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
        dv = DataValidation(type="list", formula1="lst_Ranks_All", showErrorMessage=True)
        dv.add(f"{L(cols.index('Rank') + 1)}5:{L(cols.index('Rank') + 1)}{4 + n}")
        ws.add_data_validation(dv)
        for h in ("Number of posts", "Hours FY 2027-28", "Hours FY 2028-29"):
            if h in cols:
                ci = cols.index(h) + 1
                dvn = DataValidation(type="decimal", operator="between", formula1="0", formula2="100000",
                                     showErrorMessage=True, error="Enter a number of 0 or more.")
                dvn.add(f"{L(ci)}5:{L(ci)}{4 + n}")
                ws.add_data_validation(dvn)
        for h in ("Start month", "End month"):
            if h in cols:
                ci = cols.index(h) + 1
                dvd = DataValidation(type="date", operator="between", formula1="DATE(2026,9,1)",
                                     formula2="DATE(2029,3,1)", showErrorMessage=True,
                                     error="Enter the first day of a month between Sep 2026 and Mar 2029.")
                dvd.add(f"{L(ci)}5:{L(ci)}{4 + n}")
                ws.add_data_validation(dvd)
                for rr in range(5, 5 + n):
                    ws.cell(rr, ci).number_format = "mmm-yyyy"
        ws.protection.sheet = True


def build(path):
    wb = Workbook()
    _cover(wb)
    _guidance(wb)
    _collection(wb, "MD")
    _collection(wb, "WT")
    _optional_tables(wb)
    _lists_sheet(wb)
    wb.properties.title = "EURO 2028 Resource Forecast Collection"
    wb.properties.keywords = f"E28-COLLECTION;v{C.TEMPLATE_VERSION}"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(__file__), "..", "output", f"E28_Collection_Template_v{C.TEMPLATE_VERSION}.xlsx")
    print(build(os.path.abspath(out)))
