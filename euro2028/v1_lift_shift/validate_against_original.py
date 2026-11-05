"""
Proves the lifted-and-shifted calculations reproduce the original workbook.

1. Reads the force figures and fixture codes from an ORIGINAL completed workbook.
2. Enters them into the matching rows of the new v1 workbook (a test copy).
3. Recalculates with LibreOffice and compares, line by line and date by date,
   the Demand Sheet, Cost Sheet and WT pages with the values the original
   workbook calculated.

Usage: python3 validate_against_original.py <original.xlsx> <new v1 workbook.xlsx> <work dir> [--super-high]

--super-high  enters every High figure and High fixture code under SUPER HIGH
              instead, and checks the Super High results equal the original
              workbook's High results (Super High uses the same logic).
The original is read locally only; nothing from it is written to the repository.
"""
import os
import re
import shutil
import subprocess
import sys
import warnings
from collections import Counter

import openpyxl

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_calc_pages as B  # noqa: E402
import layout_v1 as LY  # noqa: E402
from swap_headers import PERM  # noqa: E402  original day-first column -> new risk-first column

OLD_MD, OLD_WT = "MD Resource Collection Survey", "WT Resource Collection Survey"
OLD_CFG = {10: 7, 13: 10, 19: 13, 22: 16, 28: 19, 31: 22, 37: 25, 40: 28}          # old first row -> new first row
OLD_CNT = {46: 33, 47: 34, 48: 35, 49: 36, 51: 37, 52: 38, 53: 39, 54: 40, 56: 41, 57: 42, 58: 43, 59: 44,
           61: 45, 62: 46, 63: 47, 64: 48}
OLD_WT_CFG = {9: 6, 12: 9, 18: 12, 21: 15}
OLD_WT_CNT = {27: 20, 28: 21, 30: 22, 31: 23, 33: 24, 34: 25, 36: 26, 37: 27}


def norm(s):
    s = (s or "").lower().replace("\xa0", " ")
    sup = "MA" if s.strip().startswith("mutual aid") else ("FS" if s.strip().startswith("force supplied") else None)
    s = re.sub(r"^(force supplied|mutual aid requested|mutual aid)\s*-?\s*", "", s.strip())
    return sup, re.sub(r"[^a-z0-9]", "", s)


def blocks(ws, first, last):
    out, cur = {}, None
    for r in range(first, last + 1):
        if ws.cell(r, 2).value:
            cur = ws.cell(r, 2).value
        out[r] = cur
    return out


def row_maps(old_wb, t):
    """old survey row -> new survey row, for MD and WT."""
    md = {}
    for o, n in OLD_CFG.items():
        for i in range(3):
            md[o + i] = n + i
    md.update(OLD_CNT)
    om = old_wb[OLD_MD]
    ob = blocks(om, 7, 280)
    new_idx = {}
    for r in LY.md_officer_rows():
        if r in LY.MD_ROLE_ROWS | LY.MD_RANK_DROPDOWN_ROWS | LY.MD_ADDITIONAL_ROWS:
            continue
        new_idx[(*norm(t.label(LY.MD_SHEET, r)), t.rank(LY.MD_SHEET, r))] = r
    for r in range(66, 268):
        key = (*norm(ob.get(r)), LY.norm_rank(om.cell(r, 3).value))
        if key in new_idx:
            md[r] = new_idx[key]
    wt = {}
    for o, n in OLD_WT_CFG.items():
        for i in range(3):
            wt[o + i] = n + i
    wt.update(OLD_WT_CNT)
    ow = old_wb[OLD_WT]
    obw = blocks(ow, 6, 92)
    new_w = {(*norm(t.label(LY.WT_SHEET, r)), t.rank(LY.WT_SHEET, r)): r for r in LY.wt_officer_rows()
             if r not in LY.WT_ROLE_ROWS | LY.WT_ADDITIONAL_ROWS}
    for r in range(38, 92):
        key = (*norm(obw.get(r)), LY.norm_rank(ow.cell(r, 3).value))
        if key in new_w:
            wt[r] = new_w[key]
    return md, wt


def demand_rowmap(t):
    r, out = 7, {}
    for kind, p in B.md_lines(t):
        if kind == "section":
            r += 1
            continue
        out[(p["sr"], p["cfg"])] = r
        r += 1
    return out


def wt_rowmap(t):
    r, out = 5, {}
    for kind, p in B.wt_lines(t):
        if kind == "section":
            r += 1
            continue
        out[(p["sr"], p["cfg"])] = r
        r += 1
    return out


def refs(formula, sheet):
    return [int(x) for x in re.findall(rf"'{sheet}'!\$?[A-Z]+\$?(\d+)", formula or "")]


def main(orig, new, work, super_high=False):
    def risk(s):                       # original risk index -> risk index used in the new workbook
        return 3 if (super_high and s == 2) else s
    os.makedirs(work, exist_ok=True)
    oldv = openpyxl.load_workbook(orig, data_only=True)
    oldf = openpyxl.load_workbook(orig)
    t = LY.Template(os.path.join(HERE, "source", "BLANK_Euro_2028_Resource_NEW.xlsx"))
    md_map, wt_map = row_maps(oldv, t)
    # ---- fill the new workbook (test copy) with the original's inputs and fixture codes
    wb = openpyxl.load_workbook(new)
    for old_r, new_r in md_map.items():
        for c in range(4, 40):
            nc = PERM[c] + (12 if super_high and (c - 4) % 9 // 3 == 2 else 0)
            wb[LY.MD_SHEET].cell(new_r, nc).value = oldv[OLD_MD].cell(old_r, c).value
    for old_r, new_r in wt_map.items():
        for c in range(4, 16):
            wb[LY.WT_SHEET].cell(new_r, c).value = oldv[OLD_WT].cell(old_r, c).value
    for k in range(B.CAL_DAYS):
        for s in range(3):
            wb["Demand Sheet"].cell(4, B.demand_col(k, risk(s))).value = oldv["Demand Sheet"].cell(4, 4 + 4 * k + s).value
    filled = os.path.join(work, "filled.xlsx")
    wb.save(filled)
    subprocess.run([shutil.which("soffice"), "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                    "--outdir", os.path.join(work, "recalc"), filled], check=True, capture_output=True, timeout=900)
    nv = openpyxl.load_workbook(os.path.join(work, "recalc", "filled.xlsx"), data_only=True)
    # ---- compare Match Day demand and cost
    dmap = demand_rowmap(t)
    od, oc, of = oldv["Demand Sheet"], oldv["Cost Sheet"], oldf["Demand Sheet"]
    nd, nc = nv["Demand Sheet"], nv["Cost Sheet"]
    stats, diffs = Counter(), []
    for r in range(7, 265):
        rr = refs(of.cell(r, 4).value or of.cell(r, 8).value, OLD_MD)
        if not rr:
            continue
        if len(set(rr)) == 2:
            cfg_o, cnt_o = min(rr), max(rr)
            key = (md_map.get(cnt_o), md_map.get(cfg_o))
        else:
            key = (md_map.get(rr[0]), None)
        if None in (key[0],) or key not in dmap:
            stats["md_rows_unmapped"] += 1
            continue
        stats["md_rows_compared"] += 1
        nr = dmap[key]
        for k in range(B.CAL_DAYS):
            for s in range(3):
                ov, nvv = od.cell(r, 4 + 4 * k + s).value or 0, nd.cell(nr, B.demand_col(k, risk(s))).value or 0
                stats["demand_cells"] += 1
                if abs(ov - nvv) > 1e-9:
                    stats["demand_diff"] += 1
                    diffs.append(("demand", r, nr, k, s, ov, nvv))
                ovc, nvc = oc.cell(r, 4 + 3 * k + s).value or 0, nc.cell(nr, B.cost_col(k, risk(s))).value or 0
                stats["cost_cells"] += 1
                if abs(ovc - nvc) > 0.01:
                    stats["cost_diff"] += 1
                    diffs.append(("cost", r, nr, k, s, round(ovc, 2), round(nvc, 2)))
    # ---- compare Wider Tournament
    wmap = wt_rowmap(t)
    owd, owc, owf = oldv["WT Demand Sheet"], oldv["WT Cost Sheet"], oldf["WT Demand Sheet"]
    nwd, nwc, nwt = nv["WT Demand Sheet"], nv["WT Cost Sheet"], nv["Wider Tournament Template"]
    owt = oldv["Wider Tournament Template"]
    owtf = oldf["Wider Tournament Template"]
    wtt_row_for = {}
    for tr in range(5, owtf.max_row + 1):
        m = re.search(r"'WT Cost Sheet'!\$?[A-Z]+\$?(\d+)", str(owtf.cell(tr, 4).value or ""))
        if m:
            wtt_row_for[int(m.group(1))] = tr
    for r in range(265, 345):
        rr = refs(owf.cell(r, 4).value, OLD_WT)
        if not rr:
            continue
        key = (wt_map.get(max(rr)), wt_map.get(min(rr))) if len(set(rr)) == 2 else (wt_map.get(rr[0]), None)
        if key not in wmap:
            stats["wt_rows_unmapped"] += 1
            continue
        stats["wt_rows_compared"] += 1
        nr = wmap[key]
        for i in range(4):
            for name, o_ws, n_ws, tol in (("wt_demand", owd, nwd, 1e-9), ("wt_cost", owc, nwc, 0.01)):
                ov, nvv = o_ws.cell(r, 4 + i).value or 0, n_ws.cell(nr, 4 + i).value or 0
                stats[name + "_cells"] += 1
                if abs(ov - nvv) > tol:
                    stats[name + "_diff"] += 1
                    diffs.append((name, r, nr, i, ov, nvv))
        # Wider Tournament Template total: find the original template row that reads this WT Cost Sheet row
        tr = wtt_row_for.get(r)
        if tr is None:
            continue
        ov = owt.cell(tr, 12).value or 0
        nvv = nwt.cell(nr, 12).value or 0
        stats["wt_total_cells"] += 1
        if abs(ov - nvv) > 0.01:
            stats["wt_total_diff"] += 1
            diffs.append(("wt_total", tr, nr, ov, nvv))
    print(dict(stats))
    by = Counter((d[0], d[3]) for d in diffs if d[0] in ("demand", "cost"))
    print("MD differences by (page, calendar day index):", dict(sorted(by.items())))
    for d in diffs[:12]:
        print("  ", d)
    return stats, diffs


if __name__ == "__main__":
    main(*sys.argv[1:4], super_high="--super-high" in sys.argv)
