"""
Builds the EURO 2028 force collection template (MD + WT) from catalogue.py.

Design contract with Power Query / the Python engine:
  * Column A (hidden) holds a Row Key  "<ResourceID>|<Line>"  on data rows,
    "SECTION|<collection>-<section>" on section header rows.
  * Row 6 (hidden) holds a Column Key "<Risk>|<Day>|<Measure>" on input
    columns, e.g. "H|MD-1|NUM". WT uses Day = "WT".
  * Columns D..G hold Rank/Unit, Role/Description, Supply, Location.
Nothing else in the layout is relied upon by the model.
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, Protection
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

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
FILL = lambda c: PatternFill("solid", fgColor=c)  # noqa: E731
NA_FILL = PatternFill("lightGray", fgColor=NA_GREY, bgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER_ACROSS = Alignment(horizontal="centerContinuous", vertical="center")
UNLOCK = Protection(locked=False)
LOCK = Protection(locked=True)

FIRST_INPUT_COL = 8      # H
DATA_START_ROW = 8
KEY_ROW = 6


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


def _lists_sheet(wb):
    ws = wb.create_sheet("Lists")
    lists = {}
    for code, codes in C.RANK_LISTS.items():
        lists[code] = [C.RANK_LABEL[x] for x in codes]
    lists.update(C.UNIT_LISTS)
    lists.update(C.ROLE_LISTS)
    lists["Supply"] = [s[1] for s in C.SUPPLY_TYPES]
    lists["WTOptions"] = C.WT_OPTIONS
    lists["SectionStatus"] = C.SECTION_STATUS
    lists["Forces"] = [f[0] for f in R.FORCES]
    lists["YesNo"] = ["Yes", "No"]
    for ci, (name, vals) in enumerate(lists.items(), 1):
        ws.cell(1, ci, name).font = F_BOLD
        for ri, v in enumerate(vals, 2):
            ws.cell(ri, ci, v)
        ref = f"Lists!${L(ci)}$2:${L(ci)}${len(vals) + 1}"
        wb.defined_names[f"lst_{name}"] = DefinedName(f"lst_{name}", attr_text=ref)
        ws.column_dimensions[L(ci)].width = 34
    ws.cell(1, len(lists) + 2, "TemplateVersion").font = F_BOLD
    ws.cell(2, len(lists) + 2, C.TEMPLATE_VERSION)
    ws.sheet_state = "hidden"
    ws.protection.sheet = True
    return ws


def _cover(wb):
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 60
    ws.merge_cells("B2:C2")
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
        ("cv_SubmissionNo", "Submission number", "1 for first submission; increase for each resubmission."),
        ("cv_Comments", "Comments", "Any assumptions you want the planning team to know about."),
    ]
    r = 5
    dv_force = DataValidation(type="list", formula1="=lst_Forces", allow_blank=False,
                              showErrorMessage=True, errorTitle="Force",
                              error="Choose a force from the list.")
    dv_date = DataValidation(type="date", operator="between", formula1="DATE(2026,1,1)",
                             formula2="DATE(2028,12,31)", showErrorMessage=True,
                             error="Enter a date between 2026 and 2028.")
    dv_int = DataValidation(type="whole", operator="between", formula1="1", formula2="99",
                            showErrorMessage=True, error="Enter a whole number 1-99.")
    for name, label, hint in fields:
        ws.cell(r, 2, label).font = F_BOLD
        c = ws.cell(r, 3)
        c.border = BORDER
        c.protection = UNLOCK
        c.fill = FILL("FFFFFF")
        if hint:
            c.comment = None
            ws.cell(r, 4, hint).font = F_NOTE
        wb.defined_names[name] = DefinedName(name, attr_text=f"Cover!$C${r}")
        if name == "cv_Force":
            dv_force.add(c.coordinate)
        elif name == "cv_SubmissionDate":
            dv_date.add(c.coordinate)
            c.number_format = "dd/mm/yyyy"
        elif name == "cv_SubmissionNo":
            dv_int.add(c.coordinate)
        if name == "cv_Comments":
            ws.row_dimensions[r].height = 60
            c.alignment = WRAP
        r += 1
    ws.cell(r + 1, 2, "Template version").font = F_BOLD
    ws.cell(r + 1, 3, C.TEMPLATE_VERSION)
    wb.defined_names["cv_TemplateVersion"] = DefinedName("cv_TemplateVersion", attr_text=f"Cover!$C${r + 1}")
    ws.column_dimensions["D"].width = 55
    for dv in (dv_force, dv_date, dv_int):
        ws.add_data_validation(dv)
    ws.protection.sheet = True


GUIDANCE = [
    ("h", "How to complete this collection"),
    ("p", "1. Complete the Cover sheet first (force and contact details)."),
    ("p", "2. Complete the 'MD Collection' sheet for ONE typical fixture at each risk level, and the "
          "'WT Collection' sheet for ONE typical day of the wider tournament at each risk level."),
    ("p", "3. Fixtures and opponents are not yet known. Complete every risk level: Low, Medium, High and "
          "Super High. The planning team applies your figures to fixtures once the schedule is set."),
    ("p", "4. For each section, set 'Section status' in the blue section header to Complete, Not required "
          "or In progress."),
    ("h", "What to enter"),
    ("p", "Number required - how many officers (or PSUs / serials, or shifts where the row says so)."),
    ("p", "Normal ToD - the normal tour-of-duty hours EACH officer works that day (e.g. 10)."),
    ("p", "Extra Hours - the additional hours EACH officer works beyond their normal tour of duty (e.g. 4)."),
    ("p", "Match Day -2 / -1 / +1 - the days before and after the fixture. If a date ends up being both "
          "MD+1 of one fixture and MD-2 of the next, the planning model adds both requirements together."),
    ("h", "Blank, zero and not applicable"),
    ("p", "Leave a cell BLANK if you have not yet decided. Enter 0 if you need none. Grey hatched cells are "
          "NOT APPLICABLE and are locked."),
    ("p", "A section marked 'Complete' treats remaining blanks as zero. A section marked 'Not started' or "
          "'In progress' is reported to the planning team as incomplete."),
    ("h", "PSU / serial rows"),
    ("p", "Configuration rows: enter how many of each rank make up ONE Level 2 PSU or ONE Level 3 BDU / serial, "
          "and their hours. Count rows: enter how many PSUs / serials you need. The model multiplies them."),
    ("p", "Worked example: a Level 2 PSU of 1 Inspector, 3 Sergeants and 18 Constables, each working 10 hours "
          "normal ToD + 4 extra hours, and a requirement of 6 PSUs on Match Day (High) gives 6 Inspectors, "
          "18 Sergeants and 108 Constables, 1,848 hours in total."),
    ("p", "Mutual aid PSUs are costed as a package (travel day in, match day, travel day out); enter numbers only."),
    ("h", "Role profiles and additional resources"),
    ("p", "On role-profile rows, choose the role profile in column E and the rank in column D. On Additional "
          "Resources rows, describe the resource in column E, choose Force Supplied or Mutual Aid in column F "
          "and the rank in column D. Costs are calculated from the rank you choose."),
    ("p", "Protest removal: use the dedicated Protest Removal rows; do not repeat the same officers under "
          "Additional Public Order role profiles."),
    ("h", "Validation"),
    ("p", "Numbers must be whole numbers 0-9,999. Hours must be between 0 and 24. Amber cells mean something "
          "looks incomplete (e.g. a number with no hours). Red cells mean a rank or role has not been chosen."),
    ("p", "Do not insert, delete or move rows or columns; the sheet is protected so the planning model can "
          "read it automatically."),
]


def _guidance(wb):
    ws = wb.create_sheet("Guidance")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    r = 2
    ws.cell(r, 2, "Guidance").font = Font(name="Aptos", size=18, bold=True, color=NAVY)
    r += 2
    for kind, text in GUIDANCE:
        c = ws.cell(r, 2, text)
        if kind == "h":
            c.font = Font(name="Aptos", size=12, bold=True, color=BLUE)
            r += 1
        else:
            c.font = Font(name="Aptos", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * max(1, len(text) // 115 + 1)
        r += 1
    r += 1
    ws.cell(r, 2, "Legend").font = Font(name="Aptos", size=12, bold=True, color=BLUE)
    for label, fill in [("Input cell (force supplied)", FILL(WHITE)), ("Input cell (mutual aid)", FILL(MA_GREY)),
                        ("Not applicable (locked)", NA_FILL)]:
        r += 1
        ws.cell(r, 2, label).fill = fill
        ws.cell(r, 2).border = BORDER
    ws.protection.sheet = True


def _collection(wb, coll):
    ws = wb.create_sheet(f"{coll} Collection")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    cols = input_columns(coll)
    last_in = cols[-1][0]
    tot_start = last_in + 2
    tot_cols = {}  # risk -> (num col, hours col)
    for i, (risk, label, _) in enumerate(C.RISK_LEVELS):
        tot_cols[risk] = (tot_start + i * 2, tot_start + i * 2 + 1)
    last_col = tot_start + 2 * len(C.RISK_LEVELS) - 1

    widths = {1: 16, 2: 30, 3: 46, 4: 20, 5: 34, 6: 20, 7: 17}
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w
    for c, *_ in cols:
        ws.column_dimensions[L(c)].width = 10.5
    ws.column_dimensions[L(last_in + 1)].width = 2
    for c in range(tot_start, last_col + 1):
        ws.column_dimensions[L(c)].width = 11
    ws.column_dimensions["A"].hidden = True

    # ---- Header block
    title = ("Match Day" if coll == "MD" else "Wider Tournament") + " Resource Forecast Collection"
    ws.row_dimensions[1].height = 36
    for c in range(1, last_col + 1):
        ws.cell(1, c).fill = FILL(NAVY)
        ws.cell(2, c).fill = FILL(LIGHT)
    ws["B1"] = f"EURO 2028  |  {title}"
    ws["B1"].font = F_TITLE
    ws["B2"] = '="Force: "&IF(cv_Force="","(complete the Cover sheet)",cv_Force)&"   |   Template v' + \
               C.TEMPLATE_VERSION + '   |   OFFICIAL"'
    ws["B2"].font = F_BOLD
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")

    hdr = ["Row Key", "Cell Input Instructions", "Resource", "Rank / Unit", "Role Profile / Description",
           "Supply", "Location"]
    for c, h in enumerate(hdr, 1):
        for rr in (3, 4, 5):
            ws.cell(rr, c).fill = FILL(NAVY)
            ws.cell(rr, c).border = BORDER
        cell = ws.cell(5, c, h)
        cell.font = F_HDR
        cell.alignment = CENTER
    ws.row_dimensions[5].height = 30

    risk_label = {r: lbl for r, lbl, _ in C.RISK_LEVELS}
    day_label = {d: lbl for d, lbl, *_ in C.MD_DAYS}
    day_label["WT"] = "Per day (typical tournament day)"
    meas_label = dict(C.MEASURES)
    prev_risk = prev_day = None
    for c, risk, day, meas in cols:
        tint = RISK_TINT[risk]
        top = ws.cell(3, c)
        top.fill = FILL(BLUE)
        top.font = F_HDR
        top.border = BORDER
        if risk != prev_risk:
            top.value = f"{risk_label[risk].upper()} RISK"
            prev_day = None
        top.alignment = CENTER_ACROSS
        d = ws.cell(4, c)
        d.fill = FILL(tint)
        d.font = F_BOLD
        d.border = BORDER
        if day != prev_day:
            d.value = day_label[day]
        d.alignment = CENTER_ACROSS
        m = ws.cell(5, c, meas_label[meas])
        m.fill = FILL(tint)
        m.font = F_BOLD
        m.alignment = CENTER
        m.border = BORDER
        ws.cell(KEY_ROW, c, f"{risk}|{day}|{meas}")
        prev_risk, prev_day = risk, day
    for risk, (cn, ch) in tot_cols.items():
        for c, lbl in ((cn, "Total number"), (ch, "Total hours")):
            ws.cell(3, c).fill = FILL(NAVY)
            ws.cell(4, c).fill = FILL(RISK_TINT[risk])
            ws.cell(4, c).font = F_BOLD
            ws.cell(5, c, lbl).font = F_BOLD
            ws.cell(5, c).fill = FILL(RISK_TINT[risk])
            ws.cell(5, c).alignment = CENTER
            for rr in (3, 4, 5):
                ws.cell(rr, c).border = BORDER
        ws.cell(4, cn, f"{risk_label[risk]} totals").alignment = CENTER_ACROSS
        ws.cell(4, ch).alignment = CENTER_ACROSS
        ws.cell(KEY_ROW, cn, f"TOTAL|{risk}|NUM")
        ws.cell(KEY_ROW, ch, f"TOTAL|{risk}|HRS")
    ws.cell(3, tot_start, "AUTOMATIC TOTALS (all days)" if coll == "MD" else "AUTOMATIC TOTALS (per day)")
    ws.cell(3, tot_start).font = F_HDR
    for c in range(tot_start, last_col + 1):
        ws.cell(3, c).alignment = CENTER_ACROSS
    for c in range(1, 8):
        ws.cell(KEY_ROW, c, ["ROWKEY", "INSTR", "RESOURCE", "RANK", "ROLE", "SUPPLY", "LOCATION"][c - 1])
    ws.row_dimensions[KEY_ROW].hidden = True
    ws.row_dimensions[7].height = 4

    # ---- Validations
    dv_num = DataValidation(type="whole", operator="between", formula1="0", formula2="9999",
                            showErrorMessage=True, errorTitle="Number required",
                            error="Enter a whole number between 0 and 9,999, or leave blank.")
    dv_cfg = DataValidation(type="whole", operator="between", formula1="0", formula2="60",
                            showErrorMessage=True, errorTitle="Officers per PSU / serial",
                            error="Enter a whole number between 0 and 60 (officers in ONE PSU / serial).")
    dv_hrs = DataValidation(type="decimal", operator="between", formula1="0", formula2="24",
                            showErrorMessage=True, errorTitle="Hours",
                            error="Enter hours between 0 and 24, or leave blank.")
    dv_status = DataValidation(type="list", formula1="=lst_SectionStatus", showErrorMessage=True)
    dv_supply = DataValidation(type="list", formula1="=lst_Supply", showErrorMessage=True)
    dv_opt = DataValidation(type="list", formula1="=lst_WTOptions", showErrorMessage=True)
    dv_lists = {}

    def dv_list(name):
        if name not in dv_lists:
            dv_lists[name] = DataValidation(type="list", formula1=f"=lst_{name}", showErrorMessage=True,
                                            error="Choose a value from the list.")
        return dv_lists[name]

    officer_rows, rank_req_rows, role_req_rows, all_input_rows = [], [], [], []
    r = DATA_START_ROW
    sections = [(k, t) for k, t in C.SECTIONS.items() if k[0] == coll]
    for (sc_coll, sc_code), sc_title in sections:
        # section header row
        for c in range(1, last_col + 1):
            ws.cell(r, c).fill = FILL(BLUE)
            ws.cell(r, c).border = BORDER
        ws.cell(r, 1, f"SECTION|{coll}-{sc_code}")
        ws.cell(r, 2, sc_title).font = F_HDR
        ws.cell(r, 2).alignment = CENTER_ACROSS
        for c in (3, 4, 5):
            ws.cell(r, c).alignment = CENTER_ACROSS
        ws.cell(r, 6, "Section status:").font = F_HDR
        ws.cell(r, 6).alignment = Alignment(horizontal="right", vertical="center")
        st = ws.cell(r, 7, "Not started")
        st.fill = FILL("FFFFFF")
        st.font = F_BOLD
        st.protection = UNLOCK
        dv_status.add(st.coordinate)
        ws.row_dimensions[r].height = 18
        r += 1
        for res in [x for x in C.RESOURCES if x["collection"] == coll and x["section_code"] == sc_code]:
            res_lines = list(C.lines(res))
            first = r
            base_fill = FILL(MA_GREY) if res["supply"] == "MA" else FILL(WHITE)
            for ln in res_lines:
                ws.cell(r, 1, C.row_key(res["resource_id"], ln["line"]))
                for c in range(2, last_col + 1):
                    cell = ws.cell(r, c)
                    cell.border = BORDER
                    cell.font = F_BODY
                    if c != last_in + 1:
                        cell.fill = base_fill
                ws.cell(r, 2, res["instructions"] if ln["line"] == 1 else None).alignment = WRAP
                ws.cell(r, 3, res["name"] if ln["line"] == 1 else None).alignment = WRAP
                rt = res["row_type"]
                # Rank / unit column D
                dcell = ws.cell(r, 4)
                dcell.alignment = CENTER
                if rt in ("PSU_COUNT", "MA_PSU"):
                    dcell.value = ln["label"]
                    dcell.protection = UNLOCK
                    dv_list(ln["unit_list"]).add(dcell.coordinate)
                elif ln["rank"]:
                    dcell.value = ln["label"]
                else:
                    dcell.protection = UNLOCK
                    dcell.fill = FILL("FFFFFF")
                    dv_list(res["rank_list"]).add(dcell.coordinate)
                    rank_req_rows.append(r)
                # Role / description column E
                ecell = ws.cell(r, 5)
                ecell.alignment = WRAP
                if rt == "ROLE":
                    ecell.protection = UNLOCK
                    ecell.fill = FILL("FFFFFF")
                    dv_list(res["role_list"]).add(ecell.coordinate)
                    role_req_rows.append(r)
                elif rt == "ADDITIONAL":
                    ecell.protection = UNLOCK
                    ecell.fill = FILL("FFFFFF")
                    role_req_rows.append(r)
                # Supply column F
                fcell = ws.cell(r, 6)
                fcell.alignment = CENTER
                if rt == "ADDITIONAL":
                    fcell.protection = UNLOCK
                    fcell.fill = FILL("FFFFFF")
                    dv_supply.add(fcell.coordinate)
                else:
                    fcell.value = dict(C.SUPPLY_TYPES)[res["supply"]]
                # Location column G
                gcell = ws.cell(r, 7)
                gcell.alignment = CENTER
                if res["location"] == "Option":
                    gcell.protection = UNLOCK
                    gcell.fill = FILL("FFFFFF")
                    dv_opt.add(gcell.coordinate)
                else:
                    gcell.value = res["location"]
                # Inputs
                for c, risk, day, meas in cols:
                    cell = ws.cell(r, c)
                    if C.applicable(res, meas, day):
                        cell.protection = UNLOCK
                        cell.alignment = CENTER
                        if meas == "NUM":
                            (dv_cfg if rt == "PSU_CONFIG" else dv_num).add(cell.coordinate)
                            cell.number_format = "0"
                        else:
                            dv_hrs.add(cell.coordinate)
                            cell.number_format = "0.0"
                    else:
                        cell.fill = NA_FILL
                        cell.protection = LOCK
                # Totals
                for risk, (cn, ch) in tot_cols.items():
                    rc = [x for x in cols if x[1] == risk]
                    if rt == "PSU_CONFIG":   # officers-per-unit: totals are meaningless
                        ws.cell(r, cn).fill = NA_FILL
                        ws.cell(r, ch).fill = NA_FILL
                        continue
                    nums = [L(c) + str(r) for c, _, _, m in rc if m == "NUM"]
                    ws.cell(r, cn, f"=IF(COUNT({','.join(nums)})=0,\"\",SUM({','.join(nums)}))")
                    ws.cell(r, cn).number_format = "#,##0"
                    ws.cell(r, cn).alignment = CENTER
                    if rt in ("OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL"):
                        terms = []
                        days = list(dict.fromkeys(d for _, _, d, _ in rc))
                        for d in days:
                            n_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "NUM")
                            t_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "TOD")
                            x_ = next(L(c) for c, _, dd, m in rc if dd == d and m == "XH")
                            terms.append(f"N({n_}{r})*(N({t_}{r})+N({x_}{r}))")
                        ws.cell(r, ch, f"=IF(COUNT({','.join(nums)})=0,\"\",{'+'.join(terms)})")
                        ws.cell(r, ch).number_format = "#,##0.0"
                        ws.cell(r, ch).alignment = CENTER
                    else:
                        ws.cell(r, ch).fill = NA_FILL
                if rt in ("OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL"):
                    officer_rows.append(r)
                all_input_rows.append(r)
                r += 1
            if len(res_lines) > 1:
                for c in (2, 3):
                    ws.merge_cells(start_row=first, start_column=c, end_row=r - 1, end_column=c)
                    ws.cell(first, c).alignment = WRAP
            ws.row_dimensions[first].height = max(ws.row_dimensions[first].height or 15, 15)
        r += 1  # spacer
    last_row = r - 1

    # ---- Conditional formatting
    def rows_ref(col_a, col_b, rows):
        # compress consecutive rows into ranges
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

    amber = FILL("FFD966")
    red = FILL("FF7C80")
    if officer_rows:
        first_officer = min(officer_rows)
        for c, risk, day, meas in cols:
            if meas != "NUM":
                continue
            n_, t_, x_ = L(c), L(c + 1), L(c + 2)
            sq = rows_ref(n_, x_, officer_rows)
            # number entered but no hours
            ws.conditional_formatting.add(sq, FormulaRule(
                formula=[f"AND(N(${n_}{first_officer})>0,${t_}{first_officer}=\"\",${x_}{first_officer}=\"\")"],
                fill=amber, stopIfTrue=False))
            # hours entered but no number
            ws.conditional_formatting.add(sq, FormulaRule(
                formula=[f"AND(${n_}{first_officer}=\"\",OR(${t_}{first_officer}<>\"\",${x_}{first_officer}<>\"\"))"],
                fill=amber, stopIfTrue=False))
            # more than 24 hours in a day
            ws.conditional_formatting.add(sq, FormulaRule(
                formula=[f"N(${t_}{first_officer})+N(${x_}{first_officer})>24"], fill=red, stopIfTrue=False))
    rng_in = f"${L(FIRST_INPUT_COL)}{{r}}:${L(last_in)}{{r}}"
    if rank_req_rows:
        f0 = min(rank_req_rows)
        ws.conditional_formatting.add(rows_ref("D", "D", rank_req_rows), FormulaRule(
            formula=[f"AND($D{f0}=\"\",SUM({rng_in.format(r=f0)})>0)"], fill=red))
    if role_req_rows:
        f0 = min(role_req_rows)
        ws.conditional_formatting.add(rows_ref("E", "E", role_req_rows), FormulaRule(
            formula=[f"AND($E{f0}=\"\",SUM({rng_in.format(r=f0)})>0)"], fill=red))

    for dv in [dv_num, dv_cfg, dv_hrs, dv_status, dv_supply, dv_opt, *dv_lists.values()]:
        if dv.sqref:
            ws.add_data_validation(dv)

    ws.freeze_panes = ws.cell(DATA_START_ROW, FIRST_INPUT_COL)
    ws.protection.sheet = True
    ws.protection.formatColumns = False   # allow column width changes
    ws.protection.formatRows = False
    ws.print_title_rows = "3:5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    return ws, last_row


def _optional_tables(wb):
    from openpyxl.worksheet.table import Table, TableStyleInfo
    specs = [
        ("Command Team", "tblCommandTeam",
         ["Role", "Rank", "Number (FTE)", "Start month", "End month", "Supply", "Notes"], 30,
         "Optional: planning / command team posts and the months they are required (legacy 'Command Team & Planning')."),
        ("Other Resources", "tblOtherResources",
         ["Category", "Rank", "Hours FY 2027-28", "Hours FY 2028-29", "Notes"], 30,
         "Optional: other overtime hours by category (legacy 'Other Resources')."),
    ]
    legacy_categories = ["Media Management", "Licensing", "Crime Investigation", "Intelligence Arrangements",
                         "Post Event Investigation"]
    for sheet, tname, cols, n, note in specs:
        ws = wb.create_sheet(sheet)
        ws.sheet_view.showGridLines = False
        ws["A1"] = sheet
        ws["A1"].font = Font(name="Aptos", size=16, bold=True, color=NAVY)
        ws["A2"] = note
        ws["A2"].font = F_NOTE
        for ci, h in enumerate(cols, 1):
            ws.cell(4, ci, h)
            ws.column_dimensions[L(ci)].width = 24 if ci > 1 else 34
        for rr in range(5, 5 + n):
            for ci in range(1, len(cols) + 1):
                ws.cell(rr, ci).protection = UNLOCK
        if sheet == "Other Resources":
            for i, cat in enumerate(legacy_categories):
                for j, rk in enumerate(["Sergeant", "Constable"]):
                    ws.cell(5 + i * 2 + j, 1, cat)
                    ws.cell(5 + i * 2 + j, 2, rk)
        t = Table(displayName=tname, ref=f"A4:{L(len(cols))}{4 + n}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(t)
        rank_col = cols.index("Rank") + 1
        dv = DataValidation(type="list", formula1="=lst_Ranks_All", showErrorMessage=True)
        dv.add(f"{L(rank_col)}5:{L(rank_col)}{4 + n}")
        ws.add_data_validation(dv)
        for h in ("Number (FTE)", "Hours FY 2027-28", "Hours FY 2028-29"):
            if h in cols:
                dvn = DataValidation(type="decimal", operator="between", formula1="0", formula2="100000",
                                     showErrorMessage=True, error="Enter a number of 0 or more.")
                ci = cols.index(h) + 1
                dvn.add(f"{L(ci)}5:{L(ci)}{4 + n}")
                ws.add_data_validation(dvn)
        for h in ("Start month", "End month"):
            if h in cols:
                dvd = DataValidation(type="date", operator="between", formula1="DATE(2026,1,1)",
                                     formula2="DATE(2029,3,31)", showErrorMessage=True,
                                     error="Enter the first day of a month between Jan 2026 and Mar 2029.")
                ci = cols.index(h) + 1
                dvd.add(f"{L(ci)}5:{L(ci)}{4 + n}")
                ws.add_data_validation(dvd)
                for rr in range(5, 5 + n):
                    ws.cell(rr, ci).number_format = "mmm-yyyy"
        if "Supply" in cols:
            dvs = DataValidation(type="list", formula1="=lst_Supply", showErrorMessage=True)
            ci = cols.index("Supply") + 1
            dvs.add(f"{L(ci)}5:{L(ci)}{4 + n}")
            ws.add_data_validation(dvs)
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
