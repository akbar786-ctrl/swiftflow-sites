"""
Generates the legacy calculation pages for the NEW template v1, in the same
layout and formula style as the original "Base Assumptions" workbook, re-pointed
at the new rows. Output: a workbook containing ONLY the calculation pages;
inject.py then adds them to the untouched v1 template.

Legacy logic (original workbook, verified cell by cell):
  Demand Sheet  one block per date (Low | Medium | High | Super High | Total).
                Super High is NEW (not in the original): it follows exactly the
                same logic as the other risk levels, using its own columns. A fixture is
                marked by typing its code in row 4 under the risk column of its
                date. For each risk column the IF chain picks ONE relative day:
                  standard rows      MD > MD-1 > MD-2 > MD+1
                  non-matchday PSUs  MD-1 > MD-2 > MD+1
                  matchday-only PSUs MD
                PSU officers = officers per PSU (configuration) x PSUs (count).
  Cost Calculator  per officer cost = Extra Hours x hourly rate (Baseline Costs
                column X = Jun-2028); mutual-aid PSU = package (Baseline X17).
  Cost Sheet    demand x Cost Calculator for the same relative day.
  MD Demand / Cost Template   daily totals per resource (all risk levels).
  WT Demand Sheet / WT Cost Sheet / Wider Tournament Template   per risk level;
                total = cost per day x days at that risk level.
"""
import datetime as dt
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import layout_v1 as LY  # noqa: E402

CAL_START = dt.date(2028, 6, 7)          # MD-2 of the opening fixture (9 Jun 2028)
CAL_DAYS = 34                            # to 10 Jul 2028 (MD+1 of the final, 9 Jul 2028)
NR = len(LY.SCEN)                        # Match Day risk levels: Low, Medium, High, Super High
MDS, WTS = f"'{LY.MD_SHEET}'", f"'{LY.WT_SHEET}'"
BC = "'Baseline Costs'"
NAVY, BLUE, LIGHT = "1F3864", "2F5597", "D9E1F2"
thin = Side(style="thin", color="8EA9DB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_H = Font(name="Aptos Narrow", size=10, bold=True, color="FFFFFF")
F_B = Font(name="Aptos Narrow", size=9)
F_BB = Font(name="Aptos Narrow", size=9, bold=True)

# Legacy rate cells on Baseline Costs (column X = June 2028)
FS_RATE = {"Sergeant": "$X$31", "Constable": "$X$32", "Detective Sergeant": "$X$31", "Detective Constable": "$X$32"}
MA_RATE = {"Chief Superintendent": "$X$10", "Superintendent": "$X$11", "Chief Inspector": "$X$12",
           "Inspector": "$X$13", "Detective Inspector": "$X$13", "Sergeant": "$X$14", "Detective Sergeant": "$X$14",
           "Constable": "$X$15", "Detective Constable": "$X$15"}
ADDITIONAL_RATE = "$X$14"                # legacy: additional resources at the MA Sergeant rate
PACKAGE = "$X$17"                        # legacy: mutual-aid PSU package per unit per day
RANK_LOOKUP = "'Baseline Costs'!$B$37:$D$39"   # added table for Rank (Drop Down) rows
WT_DAYS = {"Low": "'Baseline Costs'!$C$43", "Medium": "'Baseline Costs'!$C$44",
           "High": "'Baseline Costs'!$C$45", "Super High": "'Baseline Costs'!$C$46"}


def hdr(ws, r, c, v=None):
    cell = ws.cell(r, c, v)
    cell.font, cell.fill, cell.border = F_H, PatternFill("solid", fgColor=NAVY), BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell


def section(ws, r, text, last_col):
    for c in range(1, last_col + 1):
        ws.cell(r, c).fill = PatternFill("solid", fgColor=BLUE)
    ws.cell(r, 1, text).font = F_H


def labels(ws, r, b, c):
    ws.cell(r, 2, b).font = F_B
    ws.cell(r, 3, c).font = F_B
    for col in (1, 2, 3):
        ws.cell(r, col).border = BORDER


# ---------------------------------------------------------------------------
# Demand lines (shared by Demand Sheet, Cost Sheet, MD templates, Fleet)
# ---------------------------------------------------------------------------
def md_lines(t):
    """[(kind, payload)] kind = 'section' | 'line'."""
    out = []
    count_rows = sorted(LY.MD_COUNTS)
    officer = LY.md_officer_rows()
    for sec in LY.MD_SECTIONS:
        if sec == 6:                       # PSU configurations: not a demand section (legacy)
            continue
        nxt = min([s for s in LY.MD_SECTIONS if s > sec] + [10 ** 6])
        rows = [r for r in count_rows + officer if sec < r < nxt]
        if not rows:
            continue
        out.append(("section", t.section_title(LY.MD_SHEET, sec)))
        for r in rows:
            label = t.label(LY.MD_SHEET, r)
            first = (r - 1) not in rows or t.label(LY.MD_SHEET, r - 1) != label
            if r in LY.MD_COUNTS:
                sup, grp, cfg = LY.MD_COUNTS[r]
                if sup == "FS":
                    for i, rk in enumerate(["Inspector", "Sergeant", "Constable"]):
                        out.append(("line", dict(sr=r, cfg=cfg + i, group=grp, kind="PSU", cc=cfg + i,
                                                 b=label if (first and i == 0) else None,
                                                 c=f'={MDS}!$C${r}&" - {rk}"')))
                else:
                    out.append(("line", dict(sr=r, cfg=None, group=grp, kind="MAPSU", cc=r,
                                             b=label if first else None, c=f"={MDS}!$C${r}")))
            else:
                out.append(("line", dict(sr=r, cfg=None, group="STD", kind="OFFICER", cc=r,
                                         b=(label or "Additional Resources") if first else None,
                                         c=f'=IF({MDS}!$C${r}="","",{MDS}!$C${r})')))
    return out


def wt_lines(t):
    out = []
    rows_all = sorted(LY.WT_COUNTS) + LY.wt_officer_rows()
    for sec in LY.WT_SECTIONS:
        if sec == 5:
            continue
        nxt = min([s for s in LY.WT_SECTIONS if s > sec] + [10 ** 6])
        rows = [r for r in rows_all if sec < r < nxt]
        if not rows:
            continue
        out.append(("section", t.section_title(LY.WT_SHEET, sec) or t.label(LY.WT_SHEET, sec)))
        for r in rows:
            label = t.label(LY.WT_SHEET, r)
            first = (r - 1) not in rows or t.label(LY.WT_SHEET, r - 1) != label
            if r in LY.WT_COUNTS:
                sup, cfg = LY.WT_COUNTS[r]
                if sup == "FS":
                    for i, rk in enumerate(["Inspector", "Sergeant", "Constable"]):
                        out.append(("line", dict(sr=r, cfg=cfg + i, kind="PSU", cc=cfg + i,
                                                 b=label if (first and i == 0) else None,
                                                 c=f'={WTS}!$C${r}&" - {rk}"')))
                else:
                    out.append(("line", dict(sr=r, cfg=None, kind="MAPSU", cc=r, b=label if first else None,
                                             c=f"={WTS}!$C${r}")))
            else:
                out.append(("line", dict(sr=r, cfg=None, kind="OFFICER", cc=r,
                                         b=(label or "Additional Resources") if first else None,
                                         c=f'=IF({WTS}!$C${r}="","",{WTS}!$C${r})')))
    return out


# ---------------------------------------------------------------------------
# Cost Calculator (survey layout: MD rows = survey rows; WT rows = survey row + WT_CC_OFFSET)
# ---------------------------------------------------------------------------
WT_CC_OFFSET = 345
MD_XH_COLS = [L(LY.md_col(d, s, 2)) for d in ("MD-2", "MD-1", "MD", "MD+1") for s in range(NR)]
WT_XH_COLS = [L(4 + 3 * s + 2) for s in range(4)]


def rate_expr(t, sheet, r, cfg_inspector_in_stadium=False):
    """Legacy hourly rate reference for a survey row, or None (no rate)."""
    roles = LY.MD_ROLE_ROWS if sheet == LY.MD_SHEET else LY.WT_ROLE_ROWS
    adds = LY.MD_ADDITIONAL_ROWS if sheet == LY.MD_SHEET else LY.WT_ADDITIONAL_ROWS
    if r in roles or r in adds:
        return f"{BC}!{ADDITIONAL_RATE}"
    sup = t.supply(sheet, r)
    if sheet == LY.MD_SHEET and r in LY.MD_RANK_DROPDOWN_ROWS:
        col = 3 if sup == "FS" else 4
        return f"IFERROR(VLOOKUP({MDS}!$C${r},{RANK_LOOKUP},{col},FALSE),0)"
    rank = t.rank(sheet, r)
    if cfg_inspector_in_stadium and rank == "Inspector":
        return f"{BC}!$X$13"
    ref = (FS_RATE if sup == "FS" else MA_RATE).get(rank)
    return f"{BC}!{ref}" if ref else None


def cost_calculator(wb, t):
    ws = wb.create_sheet("Cost Calculator")
    ws["A1"] = "Cost Calculator - cost per officer (Extra Hours x hourly rate) or per mutual-aid unit (package)"
    ws["A1"].font = F_BB
    for c, h in ((1, "Cell Input Instructions"), (2, "Resource "), (3, "Rank")):
        hdr(ws, 3, c, h)
    days = ["Match Day -2", "Match Day -1", "Match Day", "Match Day +1"]
    for i, d in enumerate(days):
        hdr(ws, 3, 4 + NR * i, d)
        for s in range(NR):
            hdr(ws, 4, 4 + NR * i + s, f"{LY.SCEN[s]}-Number required")
    hdr(ws, 3, 4 + 4 * NR, "Automatic Total Figures Calculation")
    for s in range(NR):
        hdr(ws, 4, 4 + 4 * NR + s, f"{LY.SCEN[s]}-Number required")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 18, 50, 24

    def put_row(sheet, r, out_r, n_cols, xh_cols, survey):
        cfg_is = sheet == LY.MD_SHEET and any(r in (c, c + 1, c + 2) for c in LY.MD_IN_STADIUM_CONFIG)
        labels(ws, out_r, t.label(sheet, r) or None, f'=IF({survey}!$C${r}="","",{survey}!$C${r})')
        counts = LY.MD_COUNTS if sheet == LY.MD_SHEET else LY.WT_COUNTS
        if r in counts:
            if counts[r][0] == "MA":
                for i in range(n_cols):
                    ws.cell(out_r, 4 + i, f"=IFERROR(SUM({BC}!{PACKAGE}),0)").number_format = "#,##0.00"
            return
        rate = rate_expr(t, sheet, r, cfg_is)
        for i, col in enumerate(xh_cols):
            f = f"=IFERROR(SUM({survey}!{col}{r}*{rate}),0)" if rate else 0
            ws.cell(out_r, 4 + i, f).number_format = "#,##0.00"

    ws.cell(5, 1, "MATCH DAY").font = F_BB
    for r in sorted(LY.MD_CONFIG.keys()):
        for i in range(3):
            put_row(LY.MD_SHEET, r + i, r + i, 4 * NR, MD_XH_COLS, MDS)
    last = L(3 + 4 * NR)
    for r in sorted(LY.MD_COUNTS) + LY.md_officer_rows():
        put_row(LY.MD_SHEET, r, r, 4 * NR, MD_XH_COLS, MDS)
        if r not in LY.MD_COUNTS or LY.MD_COUNTS[r][0] == "MA":
            for s in range(NR):
                tc = 4 + 4 * NR + s
                ws.cell(r, tc, f"=SUMIF($D$4:${last}$4,{L(tc)}$4,D{r}:{last}{r})").number_format = "#,##0.00"
    top = WT_CC_OFFSET
    ws.cell(top - 2, 1, "WIDER TOURNAMENT").font = F_BB
    for s in range(4):
        hdr(ws, top - 1, 4 + s, LY.WT_SCEN[s])
    for r in sorted(LY.WT_CONFIG.keys()):
        for i in range(3):
            put_row(LY.WT_SHEET, r + i, top + r + i, 4, WT_XH_COLS, WTS)
    for r in sorted(LY.WT_COUNTS) + LY.wt_officer_rows():
        put_row(LY.WT_SHEET, r, top + r, 4, WT_XH_COLS, WTS)
    ws.freeze_panes = "D5"
    return ws


# ---------------------------------------------------------------------------
# Demand Sheet and Cost Sheet (legacy IF chains)
# ---------------------------------------------------------------------------
def demand_col(k, s):
    return 4 + (NR + 1) * k + s   # D + block of 5 (Low, Medium, High, Super High, Total)


def cost_col(k, s):
    return 4 + NR * k + s         # D + block of 4 (Low, Medium, High, Super High)


def survey_num(day, s, r):
    return f"{MDS}!${L(LY.md_col(day, s, 0))}${r}"


CC_COL = {d: [L(4 + NR * i + s) for s in range(NR)] for i, d in enumerate(["MD-2", "MD-1", "MD", "MD+1"])}
CHAIN = {"STD": [(0, "MD"), (1, "MD-1"), (2, "MD-2"), (-1, "MD+1")],
         "NONMD": [(1, "MD-1"), (2, "MD-2"), (-1, "MD+1")],
         "MDONLY": [(0, "MD")]}


def chain_formula(branches):
    f = "0"
    for cond, val in reversed(branches):
        f = f"IF({cond},{val},{f})"
    return "=" + f


def calendar_pages(wb, t):
    lines = md_lines(t)
    dates = [CAL_START + dt.timedelta(i) for i in range(CAL_DAYS)]
    dem = wb.create_sheet("Demand Sheet")
    cst = wb.create_sheet("Cost Sheet")
    for ws, width, kind in ((dem, NR + 1, "D"), (cst, NR, "C")):
        for c, h in ((1, "Cell Input Instructions"), (2, "Resource "), (3, "Rank")):
            hdr(ws, 3, c, h)
        ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 14, 50, 26
        for k, d in enumerate(dates):
            c0 = 4 + width * k
            cell = hdr(ws, 3, c0, d)
            cell.number_format = "ddd dd mmm yyyy"
            for i in range(1, width):
                hdr(ws, 3, c0 + i)
            for s, name in enumerate(["L", "M", "H", "SH", "Total"][:width]):
                hdr(ws, 5, c0 + s, name)
                ws.column_dimensions[L(c0 + s)].width = 7
            for s in range(NR):
                if kind == "D":
                    m = ws.cell(4, c0 + s)
                    m.fill, m.border = PatternFill("solid", fgColor="FFF2CC"), BORDER
                else:
                    ws.cell(4, c0 + s, f"='Demand Sheet'!{L(demand_col(k, s))}4").number_format = '0;-0;;@'
    dem["A4"] = "Fixture code >"
    dem["A4"].font = F_BB
    dem["A1"] = ("Demand Sheet - type the fixture code (e.g. A1) in row 4 under Low, Medium, High or Super High on each "
                 "match date. The risk column chosen sets the fixture's risk level.")
    cst["A1"] = "Cost Sheet - fixture codes are copied from the Demand Sheet."
    for ws in (dem, cst):
        ws["A1"].font = F_BB
    r = 7
    rowmap = []
    for kind, p in lines:
        if kind == "section":
            section(dem, r, p, 4 + (NR + 1) * CAL_DAYS - 1)
            section(cst, r, p, 4 + NR * CAL_DAYS - 1)
            rowmap.append(("section", r, p))
            r += 1
            continue
        for ws in (dem, cst):
            labels(ws, r, p["b"], p["c"])
        sr, cfg, grp = p["sr"], p["cfg"], p["group"]
        for k in range(CAL_DAYS):
            for s in range(NR):
                own = demand_col(k, s)
                branches_d, branches_c = [], []
                for off, day in CHAIN[grp]:
                    kk = k + off
                    if kk < 0:
                        continue                         # before the first date (legacy edge)
                    mk = f"'Demand Sheet'!{L(demand_col(kk, s))}4>0"
                    if p["kind"] == "PSU":
                        val = f"SUM({survey_num(day, s, cfg)}*{survey_num(day, s, sr)})"
                    else:
                        val = survey_num(day, s, sr)
                    branches_d.append((mk, val))
                    mkc = f"'Cost Sheet'!${L(cost_col(kk, s))}$4>0"
                    branches_c.append((mkc, f"SUM('Demand Sheet'!{L(own)}{r}*'Cost Calculator'!{CC_COL[day][s]}{p['cc']})"))
                dem.cell(r, own, chain_formula(branches_d))
                cst.cell(r, cost_col(k, s), chain_formula(branches_c)).number_format = "#,##0"
            t0 = demand_col(k, 0)
            dem.cell(r, t0 + NR, f"=_xlfn.AGGREGATE(9,6,{L(t0)}{r}:{L(t0 + NR - 1)}{r})").font = F_BB
        rowmap.append(("line", r, p))
        r += 1
    for ws in (dem, cst):
        ws.freeze_panes = "D6"
    return rowmap, dates


def md_templates(wb, rowmap, dates):
    n = len(dates)
    for name, src in (("MD Demand Template", "Demand Sheet"), ("MD Cost Template", "Cost Sheet")):
        ws = wb.create_sheet(name)
        ws["A1"] = name + (" - officers per resource per day (all risk levels)" if "Demand" in name
                           else " - cost per resource per day (all risk levels)")
        ws["A1"].font = F_BB
        for c, h in ((1, "Cell Input Instructions"), (2, "Resource "), (3, "Rank")):
            hdr(ws, 3, c, h)
        for k, d in enumerate(dates):
            hdr(ws, 3, 4 + k, d).number_format = "dd-mmm"
            ws.column_dimensions[L(4 + k)].width = 9
        hdr(ws, 3, 4 + n, "TOTAL")
        ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 14, 50, 26
        last = 7
        for kind, r, p in rowmap:
            if kind == "section":
                section(ws, r, p, 4 + n)
                continue
            labels(ws, r, f"='{src}'!B{r}" if p["b"] else None, f"='{src}'!C{r}")
            for k in range(n):
                if src == "Demand Sheet":
                    f = f"=IFERROR('Demand Sheet'!{L(demand_col(k, NR))}{r},\"\")"
                else:
                    f = f"=IFERROR(_xlfn.AGGREGATE(9,6,'Cost Sheet'!{L(cost_col(k, 0))}{r}:{L(cost_col(k, NR - 1))}{r}),\"\")"
                ws.cell(r, 4 + k, f).number_format = "#,##0"
            ws.cell(r, 4 + n, f"=SUM(D{r}:{L(3 + n)}{r})").number_format = "#,##0"
            last = r
        tr = last + 2
        ws.cell(tr, 2, "TOTAL").font = F_BB
        for c in range(4, 5 + n):
            ws.cell(tr, c, f"=SUM({L(c)}7:{L(c)}{last})").number_format = "#,##0"
            ws.cell(tr, c).font = F_BB
        ws.freeze_panes = "D4"


def fleet(wb, rowmap, dates):
    ws = wb.create_sheet("Fleet")
    ws["A1"] = "Fleet - officers per resource per day (from the Demand Sheet)"
    ws["A1"].font = F_BB
    for c, h in ((1, "Location"), (2, "Resource "), (3, "Rank")):
        hdr(ws, 3, c, h)
    for k, d in enumerate(dates):
        hdr(ws, 3, 4 + k, d).number_format = "dd-mmm"
    for kind, r, p in rowmap:
        if kind == "section":
            ws.cell(r, 1, p).font = F_BB
            continue
        labels(ws, r, f"='Demand Sheet'!B{r}" if p["b"] else None, f"='Demand Sheet'!C{r}")
        for k in range(len(dates)):
            ws.cell(r, 4 + k, f"='Demand Sheet'!{L(demand_col(k, NR))}{r}")


def wt_pages(wb, t):
    lines = wt_lines(t)
    num = {s: L(4 + 3 * i) for i, s in enumerate(LY.WT_SCEN)}
    dem = wb.create_sheet("WT Demand Sheet")
    cst = wb.create_sheet("WT Cost Sheet")
    tpl = wb.create_sheet("Wider Tournament Template")
    for ws, heads in ((dem, LY.WT_SCEN), (cst, LY.WT_SCEN),
                      (tpl, [f"COST ON {s.upper()}" for s in LY.WT_SCEN] + [f"NUMBER OF {s.upper()}" for s in LY.WT_SCEN]
                       + ["TOTAL"])):
        for c, h in enumerate(["Cell Input Instructions", "Resource ", "Rank"] + heads, 1):
            hdr(ws, 3, c, h)
            ws.column_dimensions[L(c)].width = [14, 50, 26][c - 1] if c <= 3 else 14
    dem["A1"] = "WT Demand Sheet - officers (or mutual-aid units) per typical wider-tournament day"
    cst["A1"] = "WT Cost Sheet - cost per typical wider-tournament day"
    tpl["A1"] = ("Wider Tournament Template - cost per day x number of days at each risk level "
                 "(days set once on Baseline Costs)")
    for ws in (dem, cst, tpl):
        ws["A1"].font = F_BB
    r, first = 5, None
    for kind, p in lines:
        if kind == "section":
            for ws, n in ((dem, 7), (cst, 7), (tpl, 12)):
                section(ws, r, p, n)
            r += 1
            continue
        first = first or r
        for ws in (dem, cst, tpl):
            labels(ws, r, p["b"], p["c"])
        cc = WT_CC_OFFSET + p["cc"]
        for i, s in enumerate(LY.WT_SCEN):
            if p["kind"] == "PSU":
                f = f"=IFERROR(SUM({WTS}!{num[s]}{p['cfg']}*{WTS}!{num[s]}{p['sr']}), 0)"
            else:
                f = f"=IFERROR({WTS}!{num[s]}{p['sr']},0)"
            dem.cell(r, 4 + i, f)
            cst.cell(r, 4 + i, f"=IFERROR(SUM('WT Demand Sheet'!{L(4 + i)}{r}*'Cost Calculator'!{L(4 + i)}{cc}),0)").number_format = "#,##0"
            tpl.cell(r, 4 + i, f"='WT Cost Sheet'!{L(4 + i)}{r}").number_format = "#,##0"
            tpl.cell(r, 8 + i, f"={WT_DAYS[s]}")
        tpl.cell(r, 12, f"=IFERROR(SUM(D{r}*H{r})+SUM(E{r}*I{r})+SUM(F{r}*J{r})+SUM(G{r}*K{r}),0)").number_format = "#,##0"
        last = r
        r += 1
    for ws, cols in ((dem, range(4, 8)), (cst, range(4, 8)), (tpl, list(range(4, 8)) + [12])):
        ws.cell(last + 2, 2, "TOTAL").font = F_BB
        for c in cols:
            ws.cell(last + 2, c, f"=SUM({L(c)}{first}:{L(c)}{last})").number_format = "#,##0"
            ws.cell(last + 2, c).font = F_BB
        ws.freeze_panes = "D4"


# ---------------------------------------------------------------------------
# Static legacy pages (copied cell for cell from the original)
# ---------------------------------------------------------------------------
def static_pages(wb):
    data = json.load(open(os.path.join(HERE, "legacy_static_pages.json")))
    out = {}
    for name in ["Command Team & Planning", "Other Resources", "Baseline Costs", "GDP Deflators"]:
        d = data[name]
        ws = wb.create_sheet(name)
        for c in d["cells"]:
            cell = ws[c["ref"]]
            v = c["v"]
            if isinstance(v, dict) and "date" in v:
                v = dt.datetime.fromisoformat(v["date"])
            cell.value = v
            cell.number_format = c["nf"] or "General"
            color = c["fc"] if c["fc"] and not str(c["fc"]).startswith("theme") else None
            cell.font = Font(name="Aptos Narrow", size=c["sz"] or 11, bold=c["b"], italic=c["i"], color=color)
            if c["fill"] and not str(c["fill"]).startswith("theme"):
                cell.fill = PatternFill("solid", fgColor=c["fill"])
            elif c["fill"]:
                cell.fill = PatternFill("solid", fgColor=LIGHT)
            if c["h"] or c["wrap"]:
                cell.alignment = Alignment(horizontal=c["h"], wrap_text=c["wrap"])
            if c["bd"]:
                cell.border = BORDER
        for col, w in d["widths"].items():
            ws.column_dimensions[col].width = w
        for m in d["merges"]:
            ws.merge_cells(m)
        out[name] = ws
    # Additions on Baseline Costs (do not change any legacy cell)
    bc = out["Baseline Costs"]
    bc["B35"] = "RANK RATE LOOKUP - used by Rank (Drop Down) rows (Mounted, Dogs, Roads Policing)"
    bc["B35"].font = F_BB
    for c, h in ((2, "Rank"), (3, "Force supplied"), (4, "Mutual aid")):
        hdr(bc, 36, c, h)
    for i, (rk, fs, ma) in enumerate([("Inspector", "0", "$X$13"), ("Sergeant", "$X$31", "$X$14"),
                                      ("Constable", "$X$32", "$X$15")]):
        bc.cell(37 + i, 2, rk)
        bc.cell(37 + i, 3, 0 if fs == "0" else f"={fs}").number_format = "#,##0.00"
        bc.cell(37 + i, 4, f"={ma}").number_format = "#,##0.00"
    bc["B41"] = "WIDER TOURNAMENT - NUMBER OF DAYS AT EACH RISK LEVEL (used by Wider Tournament Template)"
    bc["B41"].font = F_BB
    hdr(bc, 42, 2, "Risk level")
    hdr(bc, 42, 3, "Days")
    for i, (s, n) in enumerate(zip(LY.WT_SCEN, [7, 11, 11, 3])):
        bc.cell(43 + i, 2, s)
        bc.cell(43 + i, 3, n).fill = PatternFill("solid", fgColor="FFF2CC")
    return out


ORDER = ["Demand Sheet", "Cost Sheet", "MD Demand Template", "MD Cost Template", "Fleet", "Cost Calculator",
         "Wider Tournament Template", "Command Team & Planning", "Other Resources", "Baseline Costs",
         "GDP Deflators", "WT Cost Sheet", "WT Demand Sheet"]
HIDDEN = {"Demand Sheet", "Cost Sheet", "Fleet", "Cost Calculator", "WT Cost Sheet", "WT Demand Sheet"}


def build(template_path, out_path):
    t = LY.Template(template_path)
    probs = LY.verify(template_path)
    if probs:
        raise SystemExit("Template does not match layout_v1:\n" + "\n".join(probs))
    wb = Workbook()
    wb.remove(wb.active)
    rowmap, dates = calendar_pages(wb, t)
    md_templates(wb, rowmap, dates)
    fleet(wb, rowmap, dates)
    cost_calculator(wb, t)
    wt_pages(wb, t)
    static_pages(wb)
    wb._sheets = [wb[n] for n in ORDER]
    for n in ORDER:
        wb[n].sheet_state = "hidden" if n in HIDDEN else "visible"
        wb[n].sheet_view.showGridLines = False
    wb.save(out_path)
    return out_path


if __name__ == "__main__":
    print(build(sys.argv[1], sys.argv[2]))
