"""
Writes the Office Script (Excel > Automate > New Script) that turns the BLANK v1
template into the full collection workbook inside Excel - the same result as
build.py and the VBA macro (make_vba.py), for organisations where VBA macros
are blocked but Office Scripts are allowed.

The row layout is embedded as data generated from the same Python definitions
build.py uses; every formula is built by loops that mirror build_calc_pages.py.

Usage: python3 make_officescript.py [out.ts]
"""
import json
import os
import sys
import warnings

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_calc_pages as B  # noqa: E402
import layout_v1 as LY  # noqa: E402
import make_vba as V  # noqa: E402  (data line builders shared with the VBA macro)

warnings.filterwarnings("ignore")
SOURCE = V.SOURCE
OUT = os.path.join(HERE, "..", "output", "Euro_2028_Build_OfficeScript.txt")


def raw(s):
    """Undo make_vba.enc (VBA literal encoding) -> plain text."""
    import re
    s = s.replace('""', '"').replace("~n~", "\n")
    return re.sub(r"~u([0-9A-F]+)~", lambda m: chr(int(m.group(1), 16)), s)


def hexcol(n):
    n = int(n)
    r, g, b = n & 255, (n >> 8) & 255, (n >> 16) & 255
    return f"#{r:02X}{g:02X}{b:02X}"


def ts_array(name, lines):
    body = ",\n".join("  " + json.dumps(raw(ln), ensure_ascii=True) for ln in lines)
    return f"const {name}: string[] = [\n{body}\n];\n"


def static_lines():
    out = {}
    for name, lines in V.static_data().items():
        conv = []
        for ln in lines:
            p = ln.split("|")
            if p[0] == "C":
                p[8] = hexcol(p[8]) if p[8] else ""
                p[9] = hexcol(p[9]) if p[9] else ""
            conv.append("|".join(p))
        out[name] = conv
    return out


def total_rows():
    ws = openpyxl.load_workbook(SOURCE)[LY.MD_SHEET]
    return [r for r in range(5, 401) if isinstance(ws.cell(r, 40).value, str) and ws.cell(r, 40).value.startswith("=")]


SCRIPT = r'''/**
 * EURO 2028 - National Policing Resource Collection
 * Builds the full collection workbook from the BLANK v1 template (Office Script).
 *
 * HOW TO USE (Excel for the web, or desktop Excel with the Automate tab)
 *  1. Open BLANK_Euro_2028_Resource_NEW.xlsx (the blank v1 template).
 *  2. Automate > New Script. Select everything in the editor and delete it.
 *  3. Paste ALL of this text, click Save script, then click Run.
 *     It takes about one to three minutes. Do not edit the workbook while it runs.
 *  4. The workbook now has the risk-level headers, Super High and all the
 *     calculation pages. Save / download it as normal.
 *
 * WHAT IT DOES
 *  - Match Day Resource: row 3 = Low | Medium | High | Super High (risk level),
 *    row 4 = Match Day -2 / -1 / Match Day / +1 x Number | Normal ToD | Extra Hours.
 *    Super High is columns AN:AY; Automatic Total Figures move to AZ:BK.
 *  - Adds the calculation pages of the original collection workbook with the
 *    original logic. Working pages are hidden as in the original.
 *  - Super High is new (not in the original) and uses the same logic and rates.
 *
 * Run it once, on the blank template. It stops if the pages already exist.
 * Classification: OFFICIAL
 */
function main(workbook: ExcelScript.Workbook) {
  const md = workbook.getWorksheet("Match Day Resource");
  if (!md || !workbook.getWorksheet("Wider Tournament Resource") || !workbook.getWorksheet("Guidance")) {
    throw new Error("Open the BLANK v1 template first - 'Match Day Resource' / 'Wider Tournament Resource' not found.");
  }
  if (workbook.getWorksheet("Demand Sheet")) {
    throw new Error("This workbook already has the calculation pages. Run the script on the blank template.");
  }
  matchDayLayout(md);
  const sh: { [name: string]: ExcelScript.Worksheet } = {};
  for (const n of PAGE_ORDER) {
    sh[n] = workbook.addWorksheet(n);
  }
  calendarPages(sh["Demand Sheet"], sh["Cost Sheet"]);
  mdTemplate(sh["MD Demand Template"], "Demand Sheet");
  mdTemplate(sh["MD Cost Template"], "Cost Sheet");
  fleetPage(sh["Fleet"]);
  costCalculatorPage(sh["Cost Calculator"]);
  wtPages(sh["WT Demand Sheet"], sh["WT Cost Sheet"], sh["Wider Tournament Template"]);
  staticPage(sh["Command Team & Planning"], ST_CTP);
  staticPage(sh["Other Resources"], ST_OTHER);
  staticPage(sh["Baseline Costs"], ST_BASELINE);
  staticPage(sh["GDP Deflators"], ST_GDP);
  baselineAdditions(sh["Baseline Costs"]);
  for (const n of PAGE_ORDER) {
    sh[n].setShowGridlines(false);
  }
  for (const n of HIDDEN) {
    sh[n].setVisibility(ExcelScript.SheetVisibility.hidden);
  }
  md.activate();
  console.log("Done. Save the workbook and send it to forces.");
}

const NR = 4;
const CAL_DAYS = 34;
const WT_CC_OFFSET = 345;
const MDS = "'Match Day Resource'";
const WTS = "'Wider Tournament Resource'";
const BCS = "'Baseline Costs'";
const NAVY = "%NAVY%";
const BLUE = "%BLUE%";
const BORDER = "#8EA9DB";
const INPUT = "#FFF2CC";
const WHITE = "#FFFFFF";
const BANNER = "%BANNER%";
const SCEN: string[] = ["Low", "Medium", "High", "Super High"];
const DAYS: string[] = ["MD-2", "MD-1", "MD", "MD+1"];
const DAY_LABEL: string[] = ["Match Day -2", "Match Day -1", "Match Day", "Match Day +1"];
const MEASURE: string[] = ["Number required", "Normal ToD", "Extra Hours"];
const PAGE_ORDER: string[] = ["Demand Sheet", "Cost Sheet", "MD Demand Template", "MD Cost Template", "Fleet",
  "Cost Calculator", "Wider Tournament Template", "Command Team & Planning", "Other Resources", "Baseline Costs",
  "GDP Deflators", "WT Cost Sheet", "WT Demand Sheet"];
const HIDDEN: string[] = ["Demand Sheet", "Cost Sheet", "Fleet", "Cost Calculator", "WT Cost Sheet", "WT Demand Sheet"];
const CHAIN: { [grp: string]: [number, string][] } = {
  STD: [[0, "MD"], [1, "MD-1"], [2, "MD-2"], [-1, "MD+1"]],
  NONMD: [[1, "MD-1"], [2, "MD-2"], [-1, "MD+1"]],
  MDONLY: [[0, "MD"]]
};

// ------------------------------------------------------------------ helpers
function colL(n: number): string {
  let s = "";
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function a1(r: number, c: number): string { return colL(c) + r; }
function box(ws: ExcelScript.Worksheet, r1: number, c1: number, r2: number, c2: number): ExcelScript.Range {
  return ws.getRange(a1(r1, c1) + ":" + a1(r2, c2));
}
function mdCol(d: string, s: number, m: number): number { return 4 + s * 12 + DAYS.indexOf(d) * 3 + m; }
function demandCol(k: number, s: number): number { return 4 + (NR + 1) * k + s; }
function costCol(k: number, s: number): number { return 4 + NR * k + s; }
function surveyNum(d: string, s: number, r: number): string { return MDS + "!$" + colL(mdCol(d, s, 0)) + "$" + r; }
function ccCol(d: string, s: number): string { return colL(4 + NR * DAYS.indexOf(d) + s); }
function dateSerial(k: number): number { return 46911 + k; }   // 46911 = 7 Jun 2028 (Excel date number)
function points(chars: number): number { return Math.round((chars * 7 + 5) * 0.75 * 100) / 100; }

function setBorders(r: ExcelScript.Range, inside: boolean) {
  const idx: ExcelScript.BorderIndex[] = [ExcelScript.BorderIndex.edgeTop, ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft, ExcelScript.BorderIndex.edgeRight];
  if (inside) {
    idx.push(ExcelScript.BorderIndex.insideVertical);
    idx.push(ExcelScript.BorderIndex.insideHorizontal);
  }
  for (const i of idx) {
    const b = r.getFormat().getRangeBorder(i);
    b.setStyle(ExcelScript.BorderLineStyle.continuous);
    b.setColor(BORDER);
  }
}
function hdr(r: ExcelScript.Range, v?: string) {
  if (v !== undefined) {
    r.setValue(v);
  }
  const f = r.getFormat();
  f.getFont().setName("Aptos Narrow");
  f.getFont().setSize(10);
  f.getFont().setBold(true);
  f.getFont().setColor(WHITE);
  f.getFill().setColor(NAVY);
  f.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  f.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  f.setWrapText(true);
  setBorders(r, true);
}
function bold9(r: ExcelScript.Range) {
  r.getFormat().getFont().setName("Aptos Narrow");
  r.getFormat().getFont().setSize(9);
  r.getFormat().getFont().setBold(true);
}
function sec(ws: ExcelScript.Worksheet, r: number, txt: string, lastCol: number) {
  box(ws, r, 1, r, lastCol).getFormat().getFill().setColor(BLUE);
  const c = ws.getRange(a1(r, 1));
  c.setValue(txt);
  c.getFormat().getFont().setName("Aptos Narrow");
  c.getFormat().getFont().setSize(10);
  c.getFormat().getFont().setBold(true);
  c.getFormat().getFont().setColor(WHITE);
}
function labels(ws: ExcelScript.Worksheet, r: number, b: string, c: string) {
  box(ws, r, 2, r, 3).setFormulas([[b, c]]);
}
function labelBlock(ws: ExcelScript.Worksheet, firstRow: number, lastRow: number) {
  const r = box(ws, firstRow, 1, lastRow, 3);
  r.getFormat().getFont().setName("Aptos Narrow");
  r.getFormat().getFont().setSize(9);
  setBorders(r, true);
}
function headerABC(ws: ExcelScript.Worksheet, first: string) {
  hdr(ws.getRange("A3"), first);
  hdr(ws.getRange("B3"), "Resource ");
  hdr(ws.getRange("C3"), "Rank");
}
function widths(ws: ExcelScript.Worksheet, a: number, b: number, c: number) {
  ws.getRange("A:A").getFormat().setColumnWidth(points(a));
  ws.getRange("B:B").getFormat().setColumnWidth(points(b));
  ws.getRange("C:C").getFormat().setColumnWidth(points(c));
}
function chainFormula(conds: string[], vals: string[]): string {
  let f = "0";
  for (let i = conds.length - 1; i >= 0; i--) {
    f = "IF(" + conds[i] + "," + vals[i] + "," + f + ")";
  }
  return "=" + f;
}

// ===================================================== 1. Match Day Resource layout
function matchDayLayout(ws: ExcelScript.Worksheet) {
  const LR = 400;
  const SPARE = 1000;
  const targ: number[] = [];
  for (let i = 0; i < 36; i++) {
    targ.push(0);
  }
  for (let d = 0; d < 4; d++) {
    for (let s = 0; s < 3; s++) {
      for (let m = 0; m < 3; m++) {
        targ[s * 12 + d * 3 + m] = 4 + d * 9 + s * 3 + m;
      }
    }
  }
  ws.getRange("D3:AV3").unmerge();
  ws.getAutoFilter().remove();
  // risk level first: rows 4 onwards of D:AM are moved (formats move with them) to a spare
  // area, then back one column at a time in risk-first order
  box(ws, 4, 4, LR, 39).moveTo(a1(4, SPARE));
  for (let p = 0; p < 36; p++) {
    const q = targ[p] - 4;
    box(ws, 4, SPARE + q, LR, SPARE + q).moveTo(a1(4, 4 + p));
  }
  // row 3 banners styled like D3 (text rewritten below)
  for (let c = 5; c <= 39; c++) {
    ws.getRange(a1(3, c)).copyFrom("D3", ExcelScript.RangeCopyType.formats);
  }
  // Super High (AN:AY) styled like High; three more total columns (BI:BK) styled like the High totals
  ws.getRange("AN:AY").insert(ExcelScript.InsertShiftDirection.right);
  ws.getRange("BI:BK").insert(ExcelScript.InsertShiftDirection.right);
  ws.getRange("AN3:AY" + LR).copyFrom("AB3:AM" + LR, ExcelScript.RangeCopyType.formats);
  ws.getRange("BI3:BK" + LR).copyFrom("BF3:BH" + LR, ExcelScript.RangeCopyType.formats);
  ws.getRange("E:BK").getFormat().setColumnWidth(points(26.85546875));
  // headers
  ws.getRange("D3:AY3").clear(ExcelScript.ClearApplyTo.contents);
  ws.getRange("BA3:BK3").clear(ExcelScript.ClearApplyTo.contents);
  const row4: string[] = [];
  for (let s = 0; s < 4; s++) {
    ws.getRange(a1(3, 4 + s * 12)).setValue(SCEN[s]);
    box(ws, 3, 4 + s * 12, 3, 15 + s * 12).merge(false);
    for (let d = 0; d < 4; d++) {
      for (let m = 0; m < 3; m++) {
        row4.push(DAY_LABEL[d] + " - " + MEASURE[m]);
      }
    }
  }
  for (let s = 0; s < 4; s++) {
    for (let m = 0; m < 3; m++) {
      row4.push(SCEN[s] + " - " + MEASURE[m]);
    }
  }
  ws.getRange("D4:BK4").setValues([row4]);
  ws.getRange("AZ3:BK3").unmerge();
  ws.getRange("AZ3:BK3").merge(false);
  ws.getRange("D3:AY3").getFormat().getFill().setColor(BANNER);
  // Automatic Total Figures: direct sums of the four day columns
  for (const r of TOTAL_ROWS) {
    const f: string[] = [];
    for (let i = 0; i < 12; i++) {
      const s = Math.floor(i / 3), m = i % 3;
      const refs: string[] = [];
      for (let d = 0; d < 4; d++) {
        refs.push(a1(r, 4 + s * 12 + d * 3 + m));
      }
      f.push("=SUM(" + refs.join(",") + ")");
    }
    box(ws, r, 52, r, 63).setFormulas([f]);
  }
  ws.getAutoFilter().apply("A3:C328");
}

// ===================================================== 2. calculation pages
function calendarPages(dem: ExcelScript.Worksheet, cst: ExcelScript.Worksheet) {
  for (let w = 0; w < 2; w++) {
    const ws = w === 0 ? dem : cst;
    const wd = w === 0 ? NR + 1 : NR;
    const last = 3 + wd * CAL_DAYS;
    headerABC(ws, "Cell Input Instructions");
    widths(ws, 14, 50, 26);
    hdr(box(ws, 3, 4, 3, last));
    hdr(box(ws, 5, 4, 5, last));
    ws.getRange(colL(4) + ":" + colL(last)).getFormat().setColumnWidth(points(7));
    const r3: (string | number)[] = [];
    const r5: string[] = [];
    const r4: string[] = [];
    for (let k = 0; k < CAL_DAYS; k++) {
      for (let s = 0; s < wd; s++) {
        r3.push(s === 0 ? dateSerial(k) : "");
        r5.push(["L", "M", "H", "SH", "Total"][s]);
        if (w === 1) {
          r4.push("='Demand Sheet'!" + colL(demandCol(k, s)) + "4");
        }
      }
    }
    box(ws, 3, 4, 3, last).setValues([r3]);
    box(ws, 3, 4, 3, last).setNumberFormat("ddd dd mmm yyyy");
    box(ws, 5, 4, 5, last).setValues([r5]);
    if (w === 1) {
      box(ws, 4, 4, 4, last).setFormulas([r4]);
      box(ws, 4, 4, 4, last).setNumberFormat("0;-0;;@");
    } else {
      for (let k = 0; k < CAL_DAYS; k++) {
        const m = box(ws, 4, demandCol(k, 0), 4, demandCol(k, NR - 1));
        m.getFormat().getFill().setColor(INPUT);
        setBorders(m, true);
      }
    }
  }
  dem.getRange("A4").setValue("Fixture code >");
  bold9(dem.getRange("A4"));
  dem.getRange("A1").setValue("Demand Sheet - type the fixture code (e.g. A1) in row 4 under Low, Medium, High or Super High on each match date. The risk column chosen sets the fixture's risk level.");
  cst.getRange("A1").setValue("Cost Sheet - fixture codes are copied from the Demand Sheet.");
  bold9(dem.getRange("A1"));
  bold9(cst.getRange("A1"));
  let r = 7;
  for (const line of MD_LINES) {
    const p = line.split("|");
    if (p[0] === "S") {
      sec(dem, r, p[1], 4 + (NR + 1) * CAL_DAYS - 1);
      sec(cst, r, p[1], 4 + NR * CAL_DAYS - 1);
    } else {
      const sr = Number(p[1]), cfg = p[2] === "" ? 0 : Number(p[2]), grp = p[3], pk = p[4], cc = Number(p[5]);
      labels(dem, r, p[6], p[7]);
      labels(cst, r, p[6], p[7]);
      const fd: string[] = [];
      const fc: string[] = [];
      for (let k = 0; k < CAL_DAYS; k++) {
        for (let s = 0; s < NR; s++) {
          const own = demandCol(k, s);
          const cd: string[] = [], vd: string[] = [], ccd: string[] = [], vc: string[] = [];
          for (const [off, day] of CHAIN[grp]) {
            const kk = k + off;
            if (kk < 0) {
              continue;
            }
            cd.push("'Demand Sheet'!" + colL(demandCol(kk, s)) + "4>0");
            vd.push(pk === "PSU" ? "SUM(" + surveyNum(day, s, cfg) + "*" + surveyNum(day, s, sr) + ")" : surveyNum(day, s, sr));
            ccd.push("'Cost Sheet'!$" + colL(costCol(kk, s)) + "$4>0");
            vc.push("SUM('Demand Sheet'!" + colL(own) + r + "*'Cost Calculator'!" + ccCol(day, s) + cc + ")");
          }
          fd.push(chainFormula(cd, vd));
          fc.push(chainFormula(ccd, vc));
        }
        const c0 = demandCol(k, 0);
        fd.push("=AGGREGATE(9,6," + a1(r, c0) + ":" + a1(r, c0 + NR - 1) + ")");
      }
      box(dem, r, 4, r, 3 + (NR + 1) * CAL_DAYS).setFormulas([fd]);
      box(cst, r, 4, r, 3 + NR * CAL_DAYS).setFormulas([fc]);
    }
    r++;
  }
  labelBlock(dem, 7, r - 1);
  labelBlock(cst, 7, r - 1);
  for (let k = 0; k < CAL_DAYS; k++) {
    box(dem, 7, demandCol(k, NR), r - 1, demandCol(k, NR)).getFormat().getFont().setBold(true);
  }
  box(cst, 7, 4, r - 1, 3 + NR * CAL_DAYS).setNumberFormat("#,##0");
  dem.getFreezePanes().freezeAt("A1:C5");
  cst.getFreezePanes().freezeAt("A1:C5");
}

function mdTemplate(ws: ExcelScript.Worksheet, src: string) {
  const isDemand = src === "Demand Sheet";
  ws.getRange("A1").setValue(ws.getName() + (isDemand ? " - officers per resource per day (all risk levels)"
    : " - cost per resource per day (all risk levels)"));
  bold9(ws.getRange("A1"));
  headerABC(ws, "Cell Input Instructions");
  hdr(box(ws, 3, 4, 3, 4 + CAL_DAYS));
  const h: (string | number)[] = [];
  for (let k = 0; k < CAL_DAYS; k++) {
    h.push(dateSerial(k));
  }
  h.push("TOTAL");
  box(ws, 3, 4, 3, 4 + CAL_DAYS).setValues([h]);
  box(ws, 3, 4, 3, 3 + CAL_DAYS).setNumberFormat("dd-mmm");
  ws.getRange(colL(4) + ":" + colL(3 + CAL_DAYS)).getFormat().setColumnWidth(points(9));
  widths(ws, 14, 50, 26);
  let r = 7, lastR = 7;
  for (const line of MD_LINES) {
    const p = line.split("|");
    if (p[0] === "S") {
      sec(ws, r, p[1], 4 + CAL_DAYS);
    } else {
      labels(ws, r, p[6] !== "" ? "='" + src + "'!B" + r : "", "='" + src + "'!C" + r);
      const f: string[] = [];
      for (let k = 0; k < CAL_DAYS; k++) {
        f.push(isDemand ? "=IFERROR('Demand Sheet'!" + colL(demandCol(k, NR)) + r + ",\"\")"
          : "=IFERROR(AGGREGATE(9,6,'Cost Sheet'!" + colL(costCol(k, 0)) + r + ":" + colL(costCol(k, NR - 1)) + r + "),\"\")");
      }
      f.push("=SUM(D" + r + ":" + colL(3 + CAL_DAYS) + r + ")");
      box(ws, r, 4, r, 4 + CAL_DAYS).setFormulas([f]);
      lastR = r;
    }
    r++;
  }
  labelBlock(ws, 7, lastR);
  box(ws, 7, 4, lastR, 4 + CAL_DAYS).setNumberFormat("#,##0");
  ws.getRange(a1(lastR + 2, 2)).setValue("TOTAL");
  const t: string[] = [];
  for (let c = 4; c <= 4 + CAL_DAYS; c++) {
    t.push("=SUM(" + colL(c) + "7:" + colL(c) + lastR + ")");
  }
  box(ws, lastR + 2, 4, lastR + 2, 4 + CAL_DAYS).setFormulas([t]);
  box(ws, lastR + 2, 4, lastR + 2, 4 + CAL_DAYS).setNumberFormat("#,##0");
  bold9(box(ws, lastR + 2, 2, lastR + 2, 4 + CAL_DAYS));
  ws.getFreezePanes().freezeAt("A1:C3");
}

function fleetPage(ws: ExcelScript.Worksheet) {
  ws.getRange("A1").setValue("Fleet - officers per resource per day (from the Demand Sheet)");
  bold9(ws.getRange("A1"));
  headerABC(ws, "Location");
  hdr(box(ws, 3, 4, 3, 3 + CAL_DAYS));
  const h: number[] = [];
  for (let k = 0; k < CAL_DAYS; k++) {
    h.push(dateSerial(k));
  }
  box(ws, 3, 4, 3, 3 + CAL_DAYS).setValues([h]);
  box(ws, 3, 4, 3, 3 + CAL_DAYS).setNumberFormat("dd-mmm");
  let r = 7;
  for (const line of MD_LINES) {
    const p = line.split("|");
    if (p[0] === "S") {
      ws.getRange(a1(r, 1)).setValue(p[1]);
      bold9(ws.getRange(a1(r, 1)));
    } else {
      labels(ws, r, p[6] !== "" ? "='Demand Sheet'!B" + r : "", "='Demand Sheet'!C" + r);
      const f: string[] = [];
      for (let k = 0; k < CAL_DAYS; k++) {
        f.push("='Demand Sheet'!" + colL(demandCol(k, NR)) + r);
      }
      box(ws, r, 4, r, 3 + CAL_DAYS).setFormulas([f]);
    }
    r++;
  }
}

function costCalculatorPage(ws: ExcelScript.Worksheet) {
  ws.getRange("A1").setValue("Cost Calculator - cost per officer (Extra Hours x hourly rate) or per mutual-aid unit (package)");
  bold9(ws.getRange("A1"));
  headerABC(ws, "Cell Input Instructions");
  for (let i = 0; i < 4; i++) {
    hdr(ws.getRange(a1(3, 4 + NR * i)), DAY_LABEL[i]);
    for (let s = 0; s < NR; s++) {
      hdr(ws.getRange(a1(4, 4 + NR * i + s)), SCEN[s] + "-Number required");
    }
  }
  hdr(ws.getRange(a1(3, 4 + 4 * NR)), "Automatic Total Figures Calculation");
  for (let s = 0; s < NR; s++) {
    hdr(ws.getRange(a1(4, 4 + 4 * NR + s)), SCEN[s] + "-Number required");
  }
  widths(ws, 18, 50, 24);
  ws.getRange("A5").setValue("MATCH DAY");
  bold9(ws.getRange("A5"));
  ws.getRange(a1(WT_CC_OFFSET - 2, 1)).setValue("WIDER TOURNAMENT");
  bold9(ws.getRange(a1(WT_CC_OFFSET - 2, 1)));
  for (let s = 0; s < 4; s++) {
    hdr(ws.getRange(a1(WT_CC_OFFSET - 1, 4 + s)), SCEN[s]);
  }
  const lastC = colL(3 + 4 * NR);
  for (const line of CC_LINES) {
    const p = line.split("|");
    const isMD = p[1] === "MD", sr = Number(p[2]), outR = Number(p[3]);
    const survey = isMD ? MDS : WTS;
    const xh: string[] = [];
    if (isMD) {
      for (let d = 0; d < 4; d++) {
        for (let s = 0; s < NR; s++) {
          xh.push(colL(4 + s * 12 + d * 3 + 2));
        }
      }
    } else {
      for (let s = 0; s < 4; s++) {
        xh.push(colL(4 + 3 * s + 2));
      }
    }
    labels(ws, outR, p[4], p[5]);
    labelBlock(ws, outR, outR);
    if (p[6] !== "SKIP") {
      const f: string[] = [];
      for (const col of xh) {
        if (p[6] === "PKG") {
          f.push("=IFERROR(SUM(" + BCS + "!$X$17),0)");
        } else if (p[6] === "RATE") {
          f.push("=IFERROR(SUM(" + survey + "!" + col + sr + "*" + p[7] + "),0)");
        } else {
          f.push("0");
        }
      }
      const rg = box(ws, outR, 4, outR, 3 + xh.length);
      rg.setFormulas([f]);
      rg.setNumberFormat("#,##0.00");
    }
    if (p[8] === "1") {
      const t: string[] = [];
      for (let s = 0; s < NR; s++) {
        const tc = 4 + 4 * NR + s;
        t.push("=SUMIF($D$4:$" + lastC + "$4," + colL(tc) + "$4,D" + outR + ":" + lastC + outR + ")");
      }
      const rg = box(ws, outR, 4 + 4 * NR, outR, 3 + 5 * NR);
      rg.setFormulas([t]);
      rg.setNumberFormat("#,##0.00");
    }
  }
  ws.getFreezePanes().freezeAt("A1:C4");
}

function wtPages(dem: ExcelScript.Worksheet, cst: ExcelScript.Worksheet, tpl: ExcelScript.Worksheet) {
  const pages: ExcelScript.Worksheet[] = [dem, cst, tpl];
  for (let w = 0; w < 3; w++) {
    const ws = pages[w];
    let heads: string[] = ["Cell Input Instructions", "Resource ", "Rank"];
    if (w < 2) {
      heads = heads.concat(SCEN);
    } else {
      heads = heads.concat(SCEN.map(s => "COST ON " + s.toUpperCase()), SCEN.map(s => "NUMBER OF " + s.toUpperCase()), ["TOTAL"]);
    }
    for (let c = 0; c < heads.length; c++) {
      hdr(ws.getRange(a1(3, c + 1)), heads[c]);
    }
    widths(ws, 14, 50, 26);
    ws.getRange("D:" + colL(heads.length)).getFormat().setColumnWidth(points(14));
  }
  dem.getRange("A1").setValue("WT Demand Sheet - officers (or mutual-aid units) per typical wider-tournament day");
  cst.getRange("A1").setValue("WT Cost Sheet - cost per typical wider-tournament day");
  tpl.getRange("A1").setValue("Wider Tournament Template - cost per day x number of days at each risk level (days set once on Baseline Costs)");
  for (const ws of pages) {
    bold9(ws.getRange("A1"));
  }
  let r = 5, firstR = 0, lastR = 0;
  for (const line of WT_LINES) {
    const p = line.split("|");
    if (p[0] === "S") {
      sec(dem, r, p[1], 7);
      sec(cst, r, p[1], 7);
      sec(tpl, r, p[1], 12);
    } else {
      if (firstR === 0) {
        firstR = r;
      }
      const sr = Number(p[1]), cfg = p[2] === "" ? 0 : Number(p[2]), cc = WT_CC_OFFSET + Number(p[4]);
      for (const ws of pages) {
        labels(ws, r, p[5], p[6]);
      }
      const fd: string[] = [], fc: string[] = [], ft: string[] = [];
      for (let i = 0; i < 4; i++) {
        const num = colL(4 + 3 * i);
        fd.push(p[3] === "PSU" ? "=IFERROR(SUM(" + WTS + "!" + num + cfg + "*" + WTS + "!" + num + sr + "), 0)"
          : "=IFERROR(" + WTS + "!" + num + sr + ",0)");
        fc.push("=IFERROR(SUM('WT Demand Sheet'!" + colL(4 + i) + r + "*'Cost Calculator'!" + colL(4 + i) + cc + "),0)");
        ft.push("='WT Cost Sheet'!" + colL(4 + i) + r);
      }
      for (let i = 0; i < 4; i++) {
        ft.push("=" + BCS + "!$C$" + (43 + i));
      }
      ft.push("=IFERROR(SUM(D" + r + "*H" + r + ")+SUM(E" + r + "*I" + r + ")+SUM(F" + r + "*J" + r + ")+SUM(G" + r + "*K" + r + "),0)");
      box(dem, r, 4, r, 7).setFormulas([fd]);
      box(cst, r, 4, r, 7).setFormulas([fc]);
      box(tpl, r, 4, r, 12).setFormulas([ft]);
      lastR = r;
    }
    r++;
  }
  for (let w = 0; w < 3; w++) {
    const ws = pages[w];
    labelBlock(ws, firstR, lastR);
    if (w > 0) {
      box(ws, firstR, 4, lastR, 7).setNumberFormat("#,##0");
    }
    if (w === 2) {
      box(ws, firstR, 12, lastR, 12).setNumberFormat("#,##0");
    }
    ws.getRange(a1(lastR + 2, 2)).setValue("TOTAL");
    bold9(ws.getRange(a1(lastR + 2, 2)));
    const cols: number[] = w < 2 ? [4, 5, 6, 7] : [4, 5, 6, 7, 12];
    for (const c of cols) {
      const cell = ws.getRange(a1(lastR + 2, c));
      cell.setFormula("=SUM(" + colL(c) + firstR + ":" + colL(c) + lastR + ")");
      cell.setNumberFormat("#,##0");
      bold9(cell);
    }
    ws.getFreezePanes().freezeAt("A1:C3");
  }
}

// ----------------------------------------------- static pages (copied from the original)
function staticPage(ws: ExcelScript.Worksheet, data: string[]) {
  for (const line of data) {
    const p = line.split("|");
    if (p[0] === "W") {
      ws.getRange(p[1] + ":" + p[1]).getFormat().setColumnWidth(points(Number(p[2])));
      continue;
    }
    if (p[0] === "M") {
      ws.getRange(p[1]).merge(false);
      continue;
    }
    const c = ws.getRange(p[1]);
    if (p[2] === "n") {
      c.setValue(Number(p[3]));
    } else if (p[2] === "d") {
      const t = Date.UTC(Number(p[3].substring(0, 4)), Number(p[3].substring(5, 7)) - 1, Number(p[3].substring(8, 10)));
      c.setValue(Math.round((t - Date.UTC(1899, 11, 30)) / 86400000));
    } else if (p[2] === "f") {
      c.setFormula(p[3]);
    } else if (p[2] === "s") {
      c.setValue("'" + p[3]);              // leading apostrophe: always stored as text, e.g. "(1)"
    }
    if (p[4] !== "General") {
      c.setNumberFormat(p[4]);
    }
    const font = c.getFormat().getFont();
    font.setName("Aptos Narrow");
    font.setSize(Number(p[7]));
    if (p[5] === "1") {
      font.setBold(true);
    }
    if (p[6] === "1") {
      font.setItalic(true);
    }
    if (p[8] !== "") {
      font.setColor(p[8]);
    }
    if (p[9] !== "") {
      c.getFormat().getFill().setColor(p[9]);
    }
    if (p[10] === "center") {
      c.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    } else if (p[10] === "left") {
      c.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    } else if (p[10] === "right") {
      c.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
    } else if (p[10] === "centerContinuous") {
      c.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.centerAcrossSelection);
    }
    if (p[11] === "1") {
      c.getFormat().setWrapText(true);
    }
    if (p[12] === "1") {
      setBorders(c, false);
    }
  }
}

function baselineAdditions(bc: ExcelScript.Worksheet) {
  bc.getRange("B35").setValue("RANK RATE LOOKUP - used by Rank (Drop Down) rows (Mounted, Dogs, Roads Policing)");
  bold9(bc.getRange("B35"));
  hdr(bc.getRange("B36"), "Rank");
  hdr(bc.getRange("C36"), "Force supplied");
  hdr(bc.getRange("D36"), "Mutual aid");
  bc.getRange("B37:D39").setFormulas([["Inspector", "0", "=$X$13"], ["Sergeant", "=$X$31", "=$X$14"],
    ["Constable", "=$X$32", "=$X$15"]]);
  bc.getRange("C37:D39").setNumberFormat("#,##0.00");
  bc.getRange("B41").setValue("WIDER TOURNAMENT - NUMBER OF DAYS AT EACH RISK LEVEL (used by Wider Tournament Template)");
  bold9(bc.getRange("B41"));
  hdr(bc.getRange("B42"), "Risk level");
  hdr(bc.getRange("C42"), "Days");
  bc.getRange("B43:C46").setValues([["Low", 7], ["Medium", 11], ["High", 11], ["Super High", 3]]);
  bc.getRange("C43:C46").getFormat().getFill().setColor(INPUT);
}

// ====================================================== DATA (generated - do not edit)
'''


def build_ts():
    t = LY.Template(SOURCE)
    probs = LY.verify(SOURCE)
    if probs:
        raise SystemExit("\n".join(probs))
    st = static_lines()
    ws = openpyxl.load_workbook(SOURCE)[LY.MD_SHEET]
    banner = ws["D3"].fill.fgColor.rgb
    text = (SCRIPT.replace("%NAVY%", "#" + B.NAVY).replace("%BLUE%", "#" + B.BLUE)
            .replace("%BANNER%", "#" + banner[-6:]))
    text += f"const TOTAL_ROWS: number[] = {json.dumps(total_rows())};\n"
    text += ts_array("MD_LINES", V.md_data(t))
    text += ts_array("WT_LINES", V.wt_data(t))
    text += ts_array("CC_LINES", V.cc_data(t))
    for var, name in (("ST_CTP", "Command Team & Planning"), ("ST_OTHER", "Other Resources"),
                      ("ST_BASELINE", "Baseline Costs"), ("ST_GDP", "GDP Deflators")):
        text += ts_array(var, st[name])
    assert " any" not in text.replace("company", "") and ": any" not in text
    return text


if __name__ == "__main__":
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)
    with open(out, "w") as f:
        f.write(build_ts())
    print(out, os.path.getsize(out), "bytes")
