"""
Builds the EURO 2028 BACKUP CALCULATION WORKBOOK.

Same pages and functions as the legacy "Base Assumptions MPSTH - Uniform"
workbook, rebuilt on the new resource structure, with the LEGACY calculation
logic (Decision D-12) implemented as uniform, generated formulas:

  Input     Cover, Match Day Resource, Wider Tournament Resource (identical to the
            force collection template - a force return can be pasted straight in),
            fixture risk per date (Demand Sheet row 4), Fleet ratios,
            Command Team & Planning posts, Other Resources hours.
  Process   Baseline Costs, GDP Deflators, Cost Calculator, Demand Sheet, Cost Sheet,
            WT Demand Template, WT Cost Sheet.
  Output    MD Demand Template, MD Cost Template, WT Cost Template, Fleet,
            Command Team & Planning, Other Resources, Summary.

Legacy rules carried exactly (see docs/02_calculation_spec.md section 5):
  * day selection per date and risk level: STD MD > MD-1 > MD-2 > MD+1,
    NONMD MD-1 > MD-2 > MD+1, MDONLY MD; risk levels summed;
  * cost = officers x Extra Hours x rate (Normal ToD not costed);
  * force-supplied Inspector and above GBP 0 except In-Stadium PSU Inspectors
    (MA Inspector rate); MA PSU package per unit per day;
  * Additional Resources: MD at MA Sergeant rate, WT at HF Constable rate;
  * WT total = cost per day x national days per risk level.
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(__file__))
import build_template as T  # noqa: E402
import catalogue as C  # noqa: E402
import reference_data as R  # noqa: E402

CAL_START = dt.date(2028, 6, 7)       # MD-2 of the opening fixture (9 Jun 2028)
CAL_DAYS = 34                         # to 10 Jul 2028 (MD+1 of the final, 9 Jul 2028)
RISKS = [r for r, _, _ in C.RISK_LEVELS]
RISK_NAMES = [n for _, n, _ in C.RISK_LEVELS]
GROUP_IDX = {"STD": 1, "NONMD": 2, "MDONLY": 3}
NAVY, BLUE, LIGHT, INPUT = "1F3864", "2F5597", "D9E1F2", "FFF2CC"
F_T = Font(name="Aptos", size=16, bold=True, color=NAVY)
F_H = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
F_B = Font(name="Aptos", size=9)
F_BB = Font(name="Aptos", size=9, bold=True)
F_N = Font(name="Aptos", size=9, italic=True, color="595959")
UNLOCK = Protection(locked=False)
SVY = {coll: f"'{name}'" for coll, name in C.SHEET_NAMES.items()}


def clean_num(ref):
    """Number required, excluding invalid entries exactly as the model does (text, negative, decimals -> 0)."""
    return f"IF(AND(N({ref})>=0,N({ref})=INT(N({ref}))),N({ref}),0)"


def clean_hours(ref):
    """Hours, excluding invalid entries exactly as the model does (text, negative, above 24 -> 0)."""
    return f"IF(AND(N({ref})>=0,N({ref})<=24),N({ref}),0)"


def fill(c):
    return PatternFill("solid", fgColor=c)


def hdr(ws, r, c, v, width=None):
    cell = ws.cell(r, c, v)
    cell.font, cell.fill = F_H, fill(BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if width:
        ws.column_dimensions[L(c)].width = width
    return cell


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def title(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = F_T
    ws["A2"] = note
    ws["A2"].font = F_N
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# Baseline Costs + GDP Deflators (legacy structure and formulas)
# ---------------------------------------------------------------------------
def gdp_deflators(wb):
    ws = wb.create_sheet("GDP Deflators")
    title(ws, "GDP Deflators", "ONS outturn / OBR forecast (as in legacy workbook, March 2026).")
    for c, h in enumerate(["Financial year", "GDP deflator index", "% change on previous year"], 2):
        hdr(ws, 4, c, h, 24)
    rows = {}
    for i, (fy, idx, pct) in enumerate(R.GDP_DEFLATORS, 5):
        ws.cell(i, 2, fy)
        ws.cell(i, 3, idx)
        ws.cell(i, 4, pct)
        rows[fy] = i
    ws.protection.sheet = True
    return rows


def baseline_costs(wb, gdp_rows):
    ws = wb.create_sheet("Baseline Costs")
    title(ws, "Baseline Costs", "Legacy rate tables and uplift method (Decision D-04). Edit the yellow cells only.")
    ranks = ["Constable", "Sergeant", "Inspector", "Chief Inspector", "Supt", "Chief Supt"]
    codes = ["PC", "SGT", "INSP", "CI", "SUPT", "CSUPT"]
    ma = {rk: (tod, ot) for t, rk, tod, ot, *_ in R.BASE_RATES if t == "MA"}
    hf = {rk: (tod, ot) for t, rk, tod, ot, *_ in R.BASE_RATES if t == "HF"}
    ann = {(t, rk): v for t, rk, v in R.ANNUAL_COSTS}

    def rate_table(top, label, table, rates):
        ws.cell(top - 1, 2, label).font = F_BB
        ws.cell(top, 2, "2025-2026").font = F_BB
        for i, rk in enumerate(ranks):
            hdr(ws, top, 3 + i, rk, 13)
        rows = [("Annual Employable Costs", lambda cd: ann[(table, cd)]),
                ("Monthly Employable Costs", None),
                ("Hourly Employable Costs", lambda cd: rates[cd][0]),
                ("Hourly Overtime Costs", lambda cd: rates[cd][1])]
        for j, (lbl, fn) in enumerate(rows, 1):
            ws.cell(top + j, 2, lbl)
            for i, cd in enumerate(codes):
                cell = ws.cell(top + j, 3 + i)
                if fn is None:
                    cell.value = f"={L(3 + i)}{top + 1}/12"
                else:
                    cell.value = fn(cd)
                    cell.fill, cell.protection = fill(INPUT), UNLOCK
                cell.number_format = "#,##0.00"

    ws.column_dimensions["B"].width = 30
    rate_table(4, "MUTUAL AID RATES", "MA", ma)            # rows 4-8
    rate_table(12, "HOST FORCE OVERTIME RATES", "HF", hf)  # rows 12-16
    # deflator index and pay awards (legacy J/K and M/N)
    hdr(ws, 4, 10, "Financial year", 12)
    hdr(ws, 4, 11, "GDP deflator index", 12)
    hdr(ws, 4, 13, "Financial year", 12)
    hdr(ws, 4, 14, "Pay award", 10)
    for j, fy in enumerate(["2025-26", "2026-27", "2027-28", "2028-29"]):
        ws.cell(5 + j, 10, fy)
        ws.cell(5 + j, 11, 1 if j == 0 else f"=$K$5*(1+('GDP Deflators'!D{gdp_rows[fy]}/100))")
        ws.cell(5 + j, 13, fy)
        pa = ws.cell(5 + j, 14, dict(R.PAY_AWARDS)[fy])
        pa.fill, pa.protection, pa.number_format = fill(INPUT), UNLOCK, "0.0%"
    # MA PSU package (legacy P/Q)
    hdr(ws, 4, 16, "Mutual aid PSU package", 22)
    hdr(ws, 4, 17, "GBP", 12)
    for j, (lbl, v) in enumerate([("IDR/TDR SUB", 31601.07), ("ADMIN (5%)", "=Q5*0.05"), ("TOTAL", "=Q5+Q6")]):
        ws.cell(5 + j, 16, lbl)
        c = ws.cell(5 + j, 17, v)
        c.number_format = "#,##0.00"
        if j == 0:
            c.fill, c.protection = fill(INPUT), UNLOCK
    ws.cell(9, 16, "Covers travel day in, match day, travel day out (D-06).").font = F_N

    # monthly uplifted rates: months Sep-2026 .. Dec-2028 in C..AD (legacy layout)
    months = []
    y, m = 2026, 9
    while (y, m) <= (2028, 12):
        months.append(dt.date(y, m, 1))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    mrow = 20
    ws.cell(mrow, 2, "Hourly rates by month (uplifted)").font = F_BB
    for i, d in enumerate(months):
        c = hdr(ws, mrow, 3 + i, d, 10)
        c.number_format = "mmm-yy"
    sep27, sep28 = 3 + months.index(dt.date(2027, 9, 1)), 3 + months.index(dt.date(2028, 9, 1))
    order = [("CSUPT", "Chief Supt", 8), ("SUPT", "Supt", 7), ("CI", "Chief Inspector", 6),
             ("INSP", "Inspector", 5), ("SGT", "Sergeant", 4), ("PC", "Constable", 3)]
    rowmap = {}

    def monthly_row(r, label, first_formula, pay_next, pay_last):
        ws.cell(r, 2, label)
        for i in range(len(months)):
            col = 3 + i
            if i == 0:
                f = first_formula
            elif col == sep27:
                f = f"={L(col - 1)}{r}+({L(col - 1)}{r}*{pay_next})"
            elif col == sep28:
                f = f"={L(col - 1)}{r}+({L(col - 1)}{r}*{pay_last})"
            else:
                f = f"={L(col - 1)}{r}"
            ws.cell(r, col, f).number_format = "#,##0.00"

    r = mrow + 1
    for cd, lbl, col in order:          # MA: rate*K(26/27) + rate*N(26/27)   (legacy C10:C15)
        src = f"{L(col)}$8"
        monthly_row(r, f"MA {lbl}", f"=({src}*$K$6)+({src}*$N$6)", "$N$7", "$N$8")
        rowmap[("MA", cd)] = r
        r += 1
    monthly_row(r, "MA PSU package", "=$Q$7+($Q$7*$N$6)", "$N$7", "$N$8")    # legacy C17
    rowmap[("MAPSU", None)] = r
    r += 2
    for cd, lbl, col in order:          # HF: OT + OT*N(26/27)                  (legacy C27:C32)
        src = f"{L(col)}$16"
        monthly_row(r, f"HF {lbl}", f"={src}+({src}*$N$6)", "$N$7", "$N$8")
        rowmap[("HF", cd)] = r
        r += 1
    last_month_col = L(3 + len(months) - 1)
    name(wb, "bc_Months", f"'Baseline Costs'!$C${mrow}:${last_month_col}${mrow}")

    # rates applied at the tournament month
    r += 2
    top = r
    ws.cell(r, 2, "RATES APPLIED (legacy rules)").font = F_BB
    ws.cell(r + 1, 2, "Rate month")
    rm = ws.cell(r + 1, 3, dt.date(2028, 6, 1))
    rm.number_format, rm.fill, rm.protection = "mmm-yyyy", fill(INPUT), UNLOCK
    name(wb, "rt_Month", f"'Baseline Costs'!$C${r + 1}")
    ws.cell(r + 2, 2, "Other Resources FY 2027-28 rate month")
    rm2 = ws.cell(r + 2, 3, dt.date(2027, 6, 1))
    rm2.number_format, rm2.fill, rm2.protection = "mmm-yyyy", fill(INPUT), UNLOCK
    name(wb, "rt_Month2728", f"'Baseline Costs'!$C${r + 2}")
    r += 4
    for c, h in enumerate(["Rank", "Force supplied (GBP/hr)", "Mutual aid (GBP/hr)", "Basis"], 2):
        hdr(ws, r, c, h, 22 if c > 2 else None)
    first = r + 1

    def at_month(row, month_name="rt_Month"):
        return f"INDEX($C${row}:${last_month_col}${row},1,MATCH({month_name},bc_Months,0))"

    rank_rows = [("Chief Superintendent", "CSUPT"), ("Superintendent", "SUPT"), ("Chief Inspector", "CI"),
                 ("Inspector", "INSP"), ("Sergeant", "SGT"), ("Constable", "PC"),
                 ("Detective Inspector", "INSP"), ("Detective Sergeant", "SGT"), ("Detective Constable", "PC"),
                 ("ACC", None)]
    for i, (lbl, cd) in enumerate(rank_rows):
        rr = first + i
        ws.cell(rr, 2, lbl)
        fs = f"={at_month(rowmap[('HF', cd)])}" if cd in ("SGT", "PC") else 0
        ma_ = f"={at_month(rowmap[('MA', cd)])}" if cd else 0
        ws.cell(rr, 3, fs).number_format = "#,##0.00"
        ws.cell(rr, 4, ma_).number_format = "#,##0.00"
        ws.cell(rr, 5, "FS Inspector and above GBP 0 (D-05); HF overtime otherwise" if cd else "No ACC rate")
    last = first + len(rank_rows) - 1
    name(wb, "rt_Ranks", f"'Baseline Costs'!$B${first}:$B${last}")
    name(wb, "rt_FS", f"'Baseline Costs'!$C${first}:$C${last}")
    name(wb, "rt_MA", f"'Baseline Costs'!$D${first}:$D${last}")
    r = last + 2
    specials = [
        ("rt_MAPSU", "Mutual aid PSU / BDU package (per unit per day)", f"={at_month(rowmap[('MAPSU', None)])}"),
        ("rt_MA_INSP", "In-Stadium PSU Inspector (legacy: MA Inspector rate)", f"={at_month(rowmap[('MA', 'INSP')])}"),
        ("rt_ADD_MD", "Additional Resources - Match Day (legacy: MA Sergeant)", f"={at_month(rowmap[('MA', 'SGT')])}"),
        ("rt_ADD_WT", "Additional Resources - Wider Tournament (legacy: HF Constable)",
         f"={at_month(rowmap[('HF', 'PC')])}"),
        ("rt_HF_SGT_2728", "HF Sergeant at FY 2027-28 rate month", f"={at_month(rowmap[('HF', 'SGT')], 'rt_Month2728')}"),
        ("rt_HF_PC_2728", "HF Constable at FY 2027-28 rate month", f"={at_month(rowmap[('HF', 'PC')], 'rt_Month2728')}"),
        ("rt_HF_SGT", "HF Sergeant at rate month", f"={at_month(rowmap[('HF', 'SGT')])}"),
        ("rt_HF_PC", "HF Constable at rate month", f"={at_month(rowmap[('HF', 'PC')])}"),
    ]
    for nm, lbl, f in specials:
        ws.cell(r, 2, lbl)
        ws.cell(r, 3, f).number_format = "#,##0.00"
        name(wb, nm, f"'Baseline Costs'!$C${r}")
        r += 1
    r += 1
    ws.cell(r, 2, "WIDER TOURNAMENT DAYS PER RISK LEVEL (national assumption, D-08)").font = F_BB
    r += 1
    for i, (risk, days, _) in enumerate(R.WT_DAY_ASSUMPTIONS):
        ws.cell(r + i, 2, RISK_NAMES[i])
        c = ws.cell(r + i, 3, days)
        c.fill, c.protection = fill(INPUT), UNLOCK
        name(wb, f"wt_Days_{risk}", f"'Baseline Costs'!$C${r + i}")
    ws.protection.sheet = True
    return rowmap, months, last_month_col


# ---------------------------------------------------------------------------
# Line definitions
# ---------------------------------------------------------------------------
def survey_rows(wb, coll):
    ws = wb[C.SHEET_NAMES[coll]]
    _, _, _, _, key_col = T.layout(coll)
    return {ws.cell(r, key_col).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, key_col).value}


def rank_ref(coll, res, ln, sr):
    """Excel expression giving the rank label of a survey line."""
    if res["row_type"] == "ROLE":
        return f"{SVY[coll]}!$A${sr}"
    if ln["rank"]:
        return f'"{C.RANK_LABEL[ln["rank"]]}"'
    return f"{SVY[coll]}!$C${sr}"


def demand_lines(coll, srows):
    """Demand Sheet lines: every survey line except PSU configurations; PSU counts expanded by rank."""
    out = []
    for res in [r for r in C.RESOURCES if r["collection"] == coll]:
        rt = res["row_type"]
        if rt == "PSU_CONFIG":
            continue
        for ln in C.lines(res):
            rk = C.row_key(res["resource_id"], ln["line"])
            sr = srows[rk]
            if rt == "PSU_COUNT":
                cfg = C.RESOURCE_BY_ID[res["psu_config"][ln["unit_level"]]]
                for cl in C.lines(cfg):
                    crk = C.row_key(cfg["resource_id"], cl["line"])
                    out.append(dict(key=f"{rk}|{cl['rank']}", res=res, label=res["name"] if ln["line"] == 1 and
                                    cl["line"] == 1 else None,
                                    rank_expr=f'{SVY[coll]}!$C${sr}&" - {C.RANK_LABEL[cl["rank"]]}"',
                                    group=res["day_group"], sr=sr, cfg_sr=srows[crk], cost_key=crk))
            else:
                out.append(dict(key=rk, res=res, label=res["name"] if ln["line"] == 1 else None,
                                rank_expr=rank_ref(coll, res, ln, sr), group=res["day_group"],
                                sr=sr, cfg_sr=None, cost_key=rk))
    return out


# ---------------------------------------------------------------------------
# Cost Calculator (survey layout): per-officer cost = Extra Hours x rate; MA PSU = package
# ---------------------------------------------------------------------------
def cost_calculator(wb, srows):
    ws = wb.create_sheet("Cost Calculator")
    title(ws, "Cost Calculator", "Per officer (or per mutual-aid unit) cost for each survey line, risk level and day. "
                                 "Legacy rule: Extra Hours x hourly rate.")
    heads = ["Row key", "Resource", "Rank / unit", "Supply", "Rate basis", "Rate (GBP/hr or per unit)"]
    for c, h in enumerate(heads, 1):
        hdr(ws, 4, c, h, [14, 46, 24, 10, 34, 14][c - 1])
    ws.column_dimensions["A"].hidden = True
    md_cols = [(RISKS[i], d) for i in range(4) for d, *_ in C.MD_DAYS]
    for j, (risk, day) in enumerate(md_cols):
        hdr(ws, 3, 7 + j, RISK_NAMES[RISKS.index(risk)], 9)
        hdr(ws, 4, 7 + j, day)
    for j in range(4):
        hdr(ws, 3, 7 + 16 + 1 + j, RISK_NAMES[j], 9)
        hdr(ws, 4, 7 + 16 + 1 + j, "WT day")
    ccrows = {}
    r = 5
    for coll in ("MD", "WT"):
        ws.cell(r, 2, "MATCH DAY" if coll == "MD" else "WIDER TOURNAMENT").font = F_BB
        r += 1
        cols = T.input_columns(coll)
        for res in [x for x in C.RESOURCES if x["collection"] == coll]:
            rt = res["row_type"]
            if rt == "PSU_COUNT":
                continue
            for ln in C.lines(res):
                rk = C.row_key(res["resource_id"], ln["line"])
                sr = srows[coll][rk]
                ws.cell(r, 1, rk)
                ws.cell(r, 2, res["name"] if ln["line"] == 1 else None)
                ws.cell(r, 3, f"={rank_ref(coll, res, ln, sr)}" if rt != "MA_PSU" else f"={SVY[coll]}!$C${sr}")
                ws.cell(r, 4, res["supply"])
                if rt == "MA_PSU":
                    basis, rate = "MA package per unit per day", \
                        f'=IF(OR($C{r}="Level 2 PSU",$C{r}="Level 3 BDU"),rt_MAPSU,0)'
                elif rt == "ADDITIONAL":
                    basis = "Legacy fixed rate (MA Sergeant)" if coll == "MD" else "Legacy fixed rate (HF Constable)"
                    rate = "=rt_ADD_MD" if coll == "MD" else "=rt_ADD_WT"
                elif res["legacy_insp_ma"] and ln["rank"] == "INSP":
                    basis, rate = "Legacy: In-Stadium Inspector at MA rate", "=rt_MA_INSP"
                else:
                    tbl = "rt_FS" if res["supply"] == "FS" else "rt_MA"
                    basis = "Force supplied rate" if res["supply"] == "FS" else "Mutual aid rate"
                    rate = f"=IFERROR(INDEX({tbl},MATCH($C{r},rt_Ranks,0)),0)"
                ws.cell(r, 5, basis)
                ws.cell(r, 6, rate).number_format = "#,##0.00"
                for c, risk, day, meas in cols:
                    if meas != "XH":
                        continue
                    out_col = 7 + (RISKS.index(risk) * 4 + C.MD_DAYS.index(next(x for x in C.MD_DAYS if x[0] == day))
                                   if coll == "MD" else 17 + RISKS.index(risk))
                    if coll == "WT":
                        out_col = 7 + 16 + 1 + RISKS.index(risk)
                    if rt == "MA_PSU":
                        f = f"=$F{r}"
                    elif C.applicable(res, "XH", day if coll == "MD" else None):
                        f = f"={clean_hours(f'{SVY[coll]}!{L(c)}{sr}')}*$F{r}"
                    else:
                        f = 0
                    ws.cell(r, out_col, f).number_format = "#,##0.00"
                ccrows[(coll, rk)] = r
                r += 1
        r += 1
    ws.freeze_panes = "G5"
    ws.protection.sheet = True
    return ccrows


# ---------------------------------------------------------------------------
# Demand Sheet + Cost Sheet (calendar)
# ---------------------------------------------------------------------------
BLOCK = 5            # 4 risk columns + total
FIRST_DATE_COL = 8   # H


def calendar_sheets(wb, srows, ccrows):
    lines = demand_lines("MD", srows["MD"])
    dates = [CAL_START + dt.timedelta(i) for i in range(CAL_DAYS)]
    dv = DataValidation(type="list", formula1='"Low,Medium,High,Super High"', showErrorMessage=True,
                        error="Choose Low, Medium, High or Super High, or leave blank (no fixture).")
    sheets = {}
    for kind in ("Demand Sheet", "Cost Sheet"):
        ws = wb.create_sheet(kind)
        sheets[kind] = ws
        if kind == "Demand Sheet":
            title(ws, "Demand Sheet", "Officers (and mutual-aid units) by date. Row 4: enter the fixture risk level on "
                                      "each date the force hosts a fixture; leave blank on other dates.")
        else:
            title(ws, "Cost Sheet", "Cost by date = Demand x Cost Calculator, for the relative day selected on the "
                                    "Demand Sheet.")
        for c, h in enumerate(["Row key", "Resource", "Rank / unit", "Group", "Survey row", "Config row", "Cost row"], 1):
            hdr(ws, 9, c, h, [14, 46, 30, 6, 6, 6, 6][c - 1])
        for c in (1, 4, 5, 6, 7):
            ws.column_dimensions[L(c)].hidden = True
        ws.cell(3, 3, "Risk index").font = F_N
        ws.cell(4, 3, "FIXTURE RISK (enter) >").font = F_BB
        ws.cell(5, 3, "Fixture risk index").font = F_N
        ws.cell(6, 3, "Day used: standard").font = F_N
        ws.cell(7, 3, "Day used: non-matchday").font = F_N
        ws.cell(8, 3, "Day used: matchday-only").font = F_N
        for k, d in enumerate(dates):
            c0 = FIRST_DATE_COL + k * BLOCK
            for i in range(BLOCK):
                col = c0 + i
                ws.column_dimensions[L(col)].width = 8 if i < 4 else 9
                dcell = hdr(ws, 2, col, d if i == 0 else None)
                dcell.number_format = "ddd dd mmm yyyy"
                dcell.alignment = Alignment(horizontal="centerContinuous", vertical="center")
                hdr(ws, 9, col, RISK_NAMES[i] if i < 4 else "Total")
                if i < 4:
                    ws.cell(3, col, i + 1)
                if kind == "Demand Sheet":
                    if i == 0:
                        inp = ws.cell(4, col)
                        inp.fill, inp.protection = fill(INPUT), UNLOCK
                        dv.add(inp.coordinate)
                    ws.cell(5, col, f'=IFERROR(MATCH(${L(c0)}$4,{{"Low","Medium","High","Super High"}},0),0)')
                    if i < 4:
                        x = L(col)
                        n1, n2, p1 = L(col + BLOCK), L(col + 2 * BLOCK), L(col - BLOCK) if k > 0 else None
                        nx1 = f"N({n1}$5)" if k + 1 < CAL_DAYS else "0"
                        nx2 = f"N({n2}$5)" if k + 2 < CAL_DAYS else "0"
                        pv1 = f"N({p1}$5)" if p1 else "0"
                        ws.cell(6, col, f"=IF({x}$5={x}$3,3,IF({nx1}={x}$3,2,IF({nx2}={x}$3,1,IF({pv1}={x}$3,4,0))))")
                        ws.cell(7, col, f"=IF({nx1}={x}$3,2,IF({nx2}={x}$3,1,IF({pv1}={x}$3,4,0)))")
                        ws.cell(8, col, f"=IF({x}$5={x}$3,3,0)")
                else:
                    for rr in (4, 5, 6, 7, 8):
                        ws.cell(rr, col, f"='Demand Sheet'!{L(col)}{rr}")
        for rr in (3, 5, 6, 7, 8):
            ws.row_dimensions[rr].hidden = True
        last_input = L(T.input_columns("MD")[-1][0])
        r = 10
        for ln in lines:
            ws.cell(r, 1, ln["key"])
            ws.cell(r, 2, ln["label"]).font = F_B
            ws.cell(r, 3, f"={ln['rank_expr']}").font = F_B
            ws.cell(r, 4, GROUP_IDX[ln["group"]])
            ws.cell(r, 5, ln["sr"])
            ws.cell(r, 6, ln["cfg_sr"])
            cr = ccrows[("MD", ln["cost_key"])]
            ws.cell(r, 7, cr)
            rng = f"{SVY['MD']}!$D${ln['sr']}:${last_input}${ln['sr']}"
            cfg = f"{SVY['MD']}!$D${ln['cfg_sr']}:${last_input}${ln['cfg_sr']}" if ln["cfg_sr"] else None
            for k in range(CAL_DAYS):
                c0 = FIRST_DATE_COL + k * BLOCK
                for i in range(4):
                    x = L(c0 + i)
                    day = f"CHOOSE($D{r},{x}$6,{x}$7,{x}$8)"
                    if kind == "Demand Sheet":
                        idx = f"({x}$3-1)*12+({day}-1)*3+1"
                        val = clean_num(f"INDEX({rng},1,{idx})") + \
                            (("*" + clean_num(f"INDEX({cfg},1,{idx})")) if cfg else "")
                        f = f"=IF({day}=0,0,{val})"
                    else:
                        f = (f"=IF({day}=0,0,'Demand Sheet'!{x}{r}*N(INDEX('Cost Calculator'!$G${cr}:$V${cr},1,"
                             f"({x}$3-1)*4+{day})))")
                    ws.cell(r, c0 + i, f)
                tcell = ws.cell(r, c0 + 4, f"=SUM({L(c0)}{r}:{L(c0 + 3)}{r})")
                tcell.font = F_BB
                if kind == "Cost Sheet":
                    for i in range(5):
                        ws.cell(r, c0 + i).number_format = "#,##0"
            r += 1
        ws.freeze_panes = ws.cell(10, FIRST_DATE_COL)
        ws.protection.sheet = True
    return lines, dates


# ---------------------------------------------------------------------------
# MD Demand / Cost templates (outputs by date), Fleet
# ---------------------------------------------------------------------------
def md_templates(wb, lines, dates):
    out = {}
    for kind, src in (("MD Demand Template", "Demand Sheet"), ("MD Cost Template", "Cost Sheet")):
        ws = wb.create_sheet(kind)
        title(ws, kind, ("Officers" if "Demand" in kind else "Cost (GBP)") +
              " by resource and date, all risk levels (from the " + src + "). Feeds the Power BI dashboard.")
        for c, h in enumerate(["Row key", "Resource", "Rank / unit"], 1):
            hdr(ws, 4, c, h, [14, 46, 30][c - 1])
        ws.column_dimensions["A"].hidden = True
        ws.cell(3, 1, "DATEKEY")
        for k, d in enumerate(dates):
            c = hdr(ws, 4, 4 + k, d, 10)
            c.number_format = "dd-mmm"
            ws.cell(3, 4 + k, d.isoformat())
        tc = 4 + len(dates)
        hdr(ws, 4, tc, "TOTAL", 12)
        ws.row_dimensions[3].hidden = True
        r = 5
        for i, ln in enumerate(lines):
            sr = 10 + i
            ws.cell(r, 1, ln["key"])
            ws.cell(r, 2, f"='{src}'!B{sr}")
            ws.cell(r, 3, f"='{src}'!C{sr}")
            for k in range(len(dates)):
                tcol = FIRST_DATE_COL + k * BLOCK + 4
                ws.cell(r, 4 + k, f"='{src}'!{L(tcol)}{sr}").number_format = "#,##0"
            ws.cell(r, tc, f"=SUM({L(4)}{r}:{L(tc - 1)}{r})").number_format = "#,##0"
            r += 1
        ws.cell(r, 2, "TOTAL").font = F_BB
        for c in range(4, tc + 1):
            ws.cell(r, c, f"=SUM({L(c)}5:{L(c)}{r - 1})").number_format = "#,##0"
            ws.cell(r, c).font = F_BB
        out[kind] = (ws, r, tc)
        ws.freeze_panes = "D5"
        ws.protection.sheet = True
    return out


def fleet(wb, lines, dates):
    ws = wb.create_sheet("Fleet")
    title(ws, "Fleet", "Officers per resource and date (from the Demand Sheet) and the equipment they require. "
                       "Enter the equipment type and the number of items per officer (e.g. 0.125 = 1 carrier per 8).")
    for c, h in enumerate(["Row key", "Resource", "Rank / unit", "Equipment", "Items per officer"], 1):
        hdr(ws, 4, c, h, [14, 40, 26, 18, 10][c - 1])
    ws.column_dimensions["A"].hidden = True
    n = len(dates)
    for k, d in enumerate(dates):
        hdr(ws, 3, 6 + k, "Officers" if k == 0 else None)
        hdr(ws, 4, 6 + k, d, 8).number_format = "dd-mmm"
        hdr(ws, 3, 7 + n + k, "Equipment required" if k == 0 else None)
        hdr(ws, 4, 7 + n + k, d, 8).number_format = "dd-mmm"
    r = 5
    for i, ln in enumerate(lines):
        sr = 10 + i
        ws.cell(r, 1, ln["key"])
        ws.cell(r, 2, f"='Demand Sheet'!B{sr}")
        ws.cell(r, 3, f"='Demand Sheet'!C{sr}")
        for c in (4, 5):
            ws.cell(r, c).fill, ws.cell(r, c).protection = fill(INPUT), UNLOCK
        for k in range(n):
            tcol = FIRST_DATE_COL + k * BLOCK + 4
            ws.cell(r, 6 + k, f"='Demand Sheet'!{L(tcol)}{sr}")
            ws.cell(r, 7 + n + k, f'=IF(N($E{r})=0,"",ROUNDUP({L(6 + k)}{r}*$E{r},0))')
        r += 1
    ws.freeze_panes = "F5"
    ws.protection.sheet = True


# ---------------------------------------------------------------------------
# Wider Tournament
# ---------------------------------------------------------------------------
def wt_sheets(wb, srows, ccrows):
    lines = demand_lines("WT", srows["WT"])
    num_col = {risk: L(T.FIRST_INPUT_COL + i * 3) for i, risk in enumerate(RISKS)}
    sheets = {}
    for kind in ("WT Demand Template", "WT Cost Sheet"):
        ws = wb.create_sheet(kind)
        sheets[kind] = ws
        title(ws, kind, "Per typical wider-tournament day, by risk level." +
              (" Officers (and mutual-aid units)." if "Demand" in kind else " Cost = Demand x Cost Calculator."))
        for c, h in enumerate(["Row key", "Resource", "Rank / unit"] + RISK_NAMES, 1):
            hdr(ws, 4, c, h, [14, 46, 30, 12, 12, 12, 12][c - 1])
        ws.column_dimensions["A"].hidden = True
        r = 5
        for ln in lines:
            ws.cell(r, 1, ln["key"])
            ws.cell(r, 2, ln["label"])
            ws.cell(r, 3, f"={ln['rank_expr']}")
            cr = ccrows[("WT", ln["cost_key"])]
            for i, risk in enumerate(RISKS):
                if kind == "WT Demand Template":
                    f = "=" + clean_num(f"{SVY['WT']}!{num_col[risk]}{ln['sr']}")
                    if ln["cfg_sr"]:
                        f += "*" + clean_num(f"{SVY['WT']}!{num_col[risk]}{ln['cfg_sr']}")
                else:
                    f = f"='WT Demand Template'!{L(4 + i)}{r}*N('Cost Calculator'!{L(7 + 17 + i)}{cr})"
                ws.cell(r, 4 + i, f).number_format = "#,##0"
            r += 1
        ws.cell(r, 2, "TOTAL").font = F_BB
        for i in range(4):
            ws.cell(r, 4 + i, f"=SUM({L(4 + i)}5:{L(4 + i)}{r - 1})").number_format = "#,##0"
        ws.freeze_panes = "D5"
        ws.protection.sheet = True
    # WT Cost Template: cost per day x national days
    ws = wb.create_sheet("WT Cost Template")
    title(ws, "WT Cost Template", "Wider tournament cost = cost per typical day x national days at each risk level "
                                  "(Baseline Costs). Feeds the Power BI dashboard.")
    heads = ["Row key", "Resource", "Rank / unit"] + [f"COST PER DAY - {n.upper()}" for n in RISK_NAMES] + \
            [f"DAYS - {n.upper()}" for n in RISK_NAMES] + ["TOTAL"]
    for c, h in enumerate(heads, 1):
        hdr(ws, 4, c, h, 14 if c > 3 else [14, 46, 30][c - 1])
    ws.column_dimensions["A"].hidden = True
    r = 5
    for j, ln in enumerate(lines):
        src_r = 5 + j
        ws.cell(r, 1, ln["key"])
        ws.cell(r, 2, f"='WT Cost Sheet'!B{src_r}")
        ws.cell(r, 3, f"='WT Cost Sheet'!C{src_r}")
        for i, risk in enumerate(RISKS):
            ws.cell(r, 4 + i, f"='WT Cost Sheet'!{L(4 + i)}{src_r}").number_format = "#,##0"
            ws.cell(r, 8 + i, f"=wt_Days_{risk}")
        ws.cell(r, 12, f"=SUMPRODUCT(D{r}:G{r},H{r}:K{r})").number_format = "#,##0"
        r += 1
    ws.cell(r, 2, "TOTAL").font = F_BB
    for c in list(range(4, 8)) + [12]:
        ws.cell(r, c, f"=SUM({L(c)}5:{L(c)}{r - 1})").number_format = "#,##0"
    ws.freeze_panes = "D5"
    ws.protection.sheet = True
    return lines, r


# ---------------------------------------------------------------------------
# Command Team & Planning, Other Resources
# ---------------------------------------------------------------------------
def command_team(wb, gdp_rows):
    ws = wb.create_sheet("Command Team & Planning")
    title(ws, "Command Team & Planning", "Monthly cost of planning / command posts (legacy method). Enter posts, "
                                         "start month and end month (last month required, inclusive - legacy rule) in the yellow cells.")
    ranks = [("Chief Superintendent", "CSUPT"), ("Superintendent", "SUPT"), ("Chief Inspector", "CI"),
             ("Inspector", "INSP"), ("Sergeant", "SGT"), ("Constable", "PC")]
    ann = {rk: v for t, rk, v in R.ANNUAL_COSTS if t == "CTP"}
    hdr(ws, 4, 2, "Inflation deflator", 22)
    hdr(ws, 4, 3, "Index", 10)
    hdr(ws, 4, 5, "Pay award", 12)
    hdr(ws, 4, 6, "Rate", 10)
    for j, fy in enumerate(["2025-26", "2026-27", "2027-28", "2028-29"]):
        ws.cell(5 + j, 2, fy)
        ws.cell(5 + j, 3, 1 if j == 0 else f"=$C$5*(1+('GDP Deflators'!D{gdp_rows[fy]}/100))")
        ws.cell(5 + j, 5, fy)
        c = ws.cell(5 + j, 6, dict(R.PAY_AWARDS)[fy])
        c.fill, c.protection, c.number_format = fill(INPUT), UNLOCK, "0.0%"
    hdr(ws, 4, 8, "2025-26 salary", 14)
    for i, (lbl, cd) in enumerate(ranks):
        hdr(ws, 4, 9 + i, lbl, 12)
        a = ws.cell(5, 9 + i, ann[cd])
        a.fill, a.protection, a.number_format = fill(INPUT), UNLOCK, "#,##0"
        ws.cell(6, 9 + i, f"={L(9 + i)}5/12").number_format = "#,##0.00"
    ws.cell(5, 8, "Annual")
    ws.cell(6, 8, "Monthly")
    months = []
    y, m = 2026, 9
    while (y, m) <= (2029, 3):
        months.append(dt.date(y, m, 1))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    # monthly rate grid (legacy rows 12-18): first month = monthly*(deflator 26/27 + pay 26/27)
    ws.cell(10, 2, "Monthly cost per post (uplifted)").font = F_BB
    hdr(ws, 11, 2, "Rank")
    for i, d in enumerate(months):
        hdr(ws, 11, 3 + i, d, 9).number_format = "mmm-yy"
    sep27, sep28 = 3 + months.index(dt.date(2027, 9, 1)), 3 + months.index(dt.date(2028, 9, 1))
    for j, (lbl, cd) in enumerate(ranks):
        r = 12 + j
        ws.cell(r, 2, lbl)
        mcol = L(9 + j)
        for i in range(len(months)):
            col = 3 + i
            if i == 0:
                f = f"=({mcol}$6*$C$6)+({mcol}$6*$F$6)"
            elif col == sep27:
                f = f"={L(col - 1)}{r}+({L(col - 1)}{r}*$F$7)"
            elif col == sep28:
                f = f"={L(col - 1)}{r}+({L(col - 1)}{r}*$F$8)"
            else:
                f = f"={L(col - 1)}{r}"
            ws.cell(r, col, f).number_format = "#,##0"
    # posts and monthly cost
    top = 20
    ws.cell(top - 1, 2, "COMMAND TEAM & PLANNING POSTS").font = F_BB
    for c, h in enumerate(["Rank", "Posts", "Start", "End"], 2):
        hdr(ws, top, c, h)
    for i, d in enumerate(months):
        hdr(ws, top, 6 + i, d, 9).number_format = "mmm-yy"
    tc = 6 + len(months)
    hdr(ws, top, tc, "TOTAL", 12)
    for j, (lbl, cd) in enumerate(ranks):
        r = top + 1 + j
        ws.cell(r, 2, lbl)
        for c in (3, 4, 5):
            ws.cell(r, c).fill, ws.cell(r, c).protection = fill(INPUT), UNLOCK
        ws.cell(r, 4).number_format = ws.cell(r, 5).number_format = "mmm-yyyy"
        for i in range(len(months)):
            mc, rc = L(6 + i), L(3 + i)
            ws.cell(r, 6 + i, f"=IF(AND(N($C{r})>0,N($D{r})>0,{mc}${top}>=$D{r},OR(N($E{r})=0,{mc}${top}<=$E{r})),"
                              f"$C{r}*{rc}{12 + j},0)").number_format = "#,##0"
        ws.cell(r, tc, f"=SUM(F{r}:{L(tc - 1)}{r})").number_format = "#,##0"
    tr = top + 1 + len(ranks)
    ws.cell(tr, 2, "TOTAL").font = F_BB
    for c in range(6, tc + 1):
        ws.cell(tr, c, f"=SUM({L(c)}{top + 1}:{L(c)}{tr - 1})").number_format = "#,##0"
    ws.protection.sheet = True
    return tr, tc


def other_resources(wb):
    ws = wb.create_sheet("Other Resources")
    title(ws, "Other Resources", "Other overtime hours by category. FY 2027-28 hours are priced at the FY 2027-28 rate "
                                 "month and FY 2028-29 hours at the tournament rate month (Baseline Costs).")
    heads = ["Category", "Rank", "Hours FY 2027-28", "Cost FY 2027-28", "Hours FY 2028-29", "Cost FY 2028-29",
             "Total cost"]
    for c, h in enumerate(heads, 1):
        hdr(ws, 4, c, h, [34, 12, 14, 14, 14, 14, 14][c - 1])
    r = 5
    for cat in R.OTHER_RESOURCE_CATEGORIES:
        for rk, n27, n28 in (("Sergeant", "rt_HF_SGT_2728", "rt_HF_SGT"), ("Constable", "rt_HF_PC_2728", "rt_HF_PC")):
            ws.cell(r, 1, cat)
            ws.cell(r, 2, rk)
            for c in (3, 5):
                ws.cell(r, c).fill, ws.cell(r, c).protection = fill(INPUT), UNLOCK
            ws.cell(r, 4, f"=N(C{r})*{n27}").number_format = "#,##0.00"
            ws.cell(r, 6, f"=N(E{r})*{n28}").number_format = "#,##0.00"
            ws.cell(r, 7, f"=D{r}+F{r}").number_format = "#,##0.00"
            r += 1
    ws.cell(r, 1, "TOTAL").font = F_BB
    for c in range(3, 8):
        ws.cell(r, c, f"=SUM({L(c)}5:{L(c)}{r - 1})").number_format = "#,##0.00"
    ws.protection.sheet = True
    return r


def summary(wb, md, wt_total_row, ctp, other_total_row):
    ws = wb.create_sheet("Summary", 2)
    title(ws, "Summary", "Totals from the calculation pages (legacy rules, RuleSet = LEGACY).")
    rows = [
        ("Match Day officers (all fixtures entered on the Demand Sheet)", f"='MD Demand Template'!"
                                                                          f"{L(md['MD Demand Template'][2])}{md['MD Demand Template'][1]}"),
        ("Match Day cost (GBP)", f"='MD Cost Template'!{L(md['MD Cost Template'][2])}{md['MD Cost Template'][1]}"),
        ("Wider Tournament cost (GBP)", f"='WT Cost Template'!L{wt_total_row}"),
        ("Command Team & Planning cost (GBP)", f"='Command Team & Planning'!{L(ctp[1])}{ctp[0]}"),
        ("Other Resources cost (GBP)", f"='Other Resources'!G{other_total_row}"),
    ]
    for i, (lbl, f) in enumerate(rows, 4):
        ws.cell(i, 2, lbl).font = F_BB
        ws.cell(i, 3, f).number_format = "#,##0"
    ws.cell(9, 2, "TOTAL (GBP)").font = F_BB
    ws.cell(9, 3, "=SUM(C5:C8)").number_format = "#,##0"
    ws.cell(9, 3).font = F_BB
    ws.cell(11, 2, "Fixtures entered on the Demand Sheet").font = F_BB
    ws.cell(11, 3, f"=COUNTA('Demand Sheet'!$H$4:${L(FIRST_DATE_COL + CAL_DAYS * BLOCK)}$4)")
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 18
    ws.protection.sheet = True


def build(path):
    wb = Workbook()
    T._cover(wb)
    T._guidance(wb)
    T._collection(wb, "MD")
    T._collection(wb, "WT")
    T._lists_sheet(wb)
    gdp = gdp_deflators(wb)
    baseline_costs(wb, gdp)
    srows = {coll: survey_rows(wb, coll) for coll in ("MD", "WT")}
    cc = cost_calculator(wb, srows)
    lines, dates = calendar_sheets(wb, srows, cc)
    md = md_templates(wb, lines, dates)
    fleet(wb, lines, dates)
    _, wt_total = wt_sheets(wb, srows, cc)
    ctp = command_team(wb, gdp)
    orr = other_resources(wb)
    summary(wb, md, wt_total, ctp, orr)
    wb["Guidance"]["B3"] = ("BACKUP CALCULATION WORKBOOK: complete the input pages (Cover, Match Day Resource, "
                            "Wider Tournament Resource, Demand Sheet row 4, Fleet, Command Team & Planning, Other "
                            "Resources). All other pages calculate automatically using the legacy rules.")
    wb["Guidance"]["B3"].font = Font(name="Aptos", size=10, bold=True, color="C00000")
    wb.properties.title = "EURO 2028 Backup Calculation Workbook"
    wb.properties.keywords = f"E28-CALC;v{C.TEMPLATE_VERSION}"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(__file__), "..", "output", f"E28_Backup_Calculation_Workbook_v{C.TEMPLATE_VERSION}.xlsx")
    print(build(os.path.abspath(out)))
