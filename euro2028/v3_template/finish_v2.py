"""
Finishes the national team's 'Resource Collection Template V2' without re-saving it through a
spreadsheet library (so banners, logos, drop-downs and comments are kept byte for byte):

  * Guidance tab replaced with the finished guidance (general, or MPS version with Pan London)
  * High columns recoloured orange, Super High red (headers and cell tints), totals headers fixed
  * Hatched + locked cells by the agreed rules, with entry blocked by validation:
      - Match Day only lines: Match Day -2 / -1 / +1 hatched; Non-Match Day lines: Match Day hatched
        (location-based public order lines - In Stadium, Stadium Vicinity, Town Centre - are Match Day only)
      - PSU / BDU number lines: hours hatched (hours are on the configuration lines)
      - In Stadium (all ranks, both origins): Chargeable hatched, Opportunity Cost open
      - Mutual Aid (not In Stadium): Opportunity Cost hatched
      - Force Supplied Inspectors (not In Stadium): Chargeable hatched
      - Additional Resources: the same origin / rank rules applied automatically (conditional + validation)
  * Missing-hours check: an open hours cell turns red when officers are entered that day and it is blank
  * All other input cells unlocked (ready for Review > Protect Sheet)

Outputs two files: general (no Pan London tab) and MPS (with Pan London).

Usage: python3 finish_v2.py <Resource_Collection_Template_V2.xlsx> <out dir>
"""
import copy
import io
import os
import re
import sys
import warnings
import zipfile

import openpyxl
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "..", "v1_lift_shift"))
import inject  # noqa: E402
import make_guidance  # noqa: E402

warnings.filterwarnings("ignore")
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
N = "{%s}" % NS
HATCH = ("lightGray", "FF808080", "FFD9D9D9")
HIGH_T = ["FFFEF4EA", "FFFDEBDA", "FFFCE2CA"]
SH_T = ["FFFFF3F3", "FFFFE8E8", "FFFFDDDD"]
PINK = {"FFFCF6FC": 0, "FFFCF2FB": 1, "FFFAECF9": 2}
HEAD3 = {"FFF5FBF3": "FFC8E6C9", "FFFEF5F0": "FFF0DC8B", "FFFCF2FB": "FFF8C471", "FFEBF1FA": "FFEF9A9A"}


def col_idx(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n, int(m.group(2))


def L(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def rects(cells):
    """set of (r, c) -> list of 'A1:B2' covering exactly those cells."""
    rows = {}
    for r, c in cells:
        rows.setdefault(r, set()).add(c)
    out, prev, start = [], None, None
    keys = sorted(rows)

    def emit(r0, r1, cols):
        cs = sorted(cols)
        i = 0
        while i < len(cs):
            j = i
            while j + 1 < len(cs) and cs[j + 1] == cs[j] + 1:
                j += 1
            out.append(f"{L(cs[i])}{r0}:{L(cs[j])}{r1}" if (r0, cs[i]) != (r1, cs[j]) else f"{L(cs[i])}{r0}")
            i = j + 1
    for r in keys:
        if prev is not None and r == prev[0] + 1 and rows[r] == prev[1]:
            prev = (r, prev[1])
            continue
        if prev is not None:
            emit(start, prev[0], prev[1])
        start, prev = r, (r, rows[r])
    if prev is not None:
        emit(start, prev[0], prev[1])
    return out


# ------------------------------------------------------------------ styles
class Styles:
    def __init__(self, xml):
        self.root = etree.fromstring(xml.encode() if isinstance(xml, str) else xml)
        self.fills = self.root.find(N + "fills")
        self.xfs = self.root.find(N + "cellXfs")
        self.dxfs = self.root.find(N + "dxfs")
        if self.dxfs is None:
            self.dxfs = etree.SubElement(self.root, N + "dxfs")
        self.fill_cache, self.xf_cache = {}, {}

    def fill_rgb(self, xf_index):
        xf = self.xfs[xf_index]
        f = self.fills[int(xf.get("fillId", 0))]
        pf = f.find(N + "patternFill")
        if pf is None:
            return None
        fg = pf.find(N + "fgColor")
        return (pf.get("patternType"), fg.get("rgb") if fg is not None else None)

    def add_fill(self, pattern, fg, bg=None):
        key = (pattern, fg, bg)
        if key not in self.fill_cache:
            f = etree.SubElement(self.fills, N + "fill")
            pf = etree.SubElement(f, N + "patternFill", patternType=pattern)
            etree.SubElement(pf, N + "fgColor", rgb=fg)
            etree.SubElement(pf, N + "bgColor", **({"rgb": bg} if bg else {"indexed": "64"}))
            self.fills.set("count", str(len(self.fills)))
            self.fill_cache[key] = len(self.fills) - 1
        return self.fill_cache[key]

    def variant(self, xf_index, fill_id=None, locked=None, font_white=None):
        key = (xf_index, fill_id, locked)
        if key in self.xf_cache:
            return self.xf_cache[key]
        xf = copy.deepcopy(self.xfs[xf_index])
        if fill_id is not None:
            xf.set("fillId", str(fill_id))
            xf.set("applyFill", "1")
        if locked is not None:
            prot = xf.find(N + "protection")
            if prot is None:
                prot = etree.SubElement(xf, N + "protection")
            prot.set("locked", "1" if locked else "0")
            xf.set("applyProtection", "1")
        self.xfs.append(xf)
        self.xfs.set("count", str(len(self.xfs)))
        self.xf_cache[key] = len(self.xfs) - 1
        return self.xf_cache[key]

    def add_dxf(self, fill_pattern, fg, bg=None, font_color=None, bold=False):
        d = etree.SubElement(self.dxfs, N + "dxf")
        if font_color:
            f = etree.SubElement(d, N + "font")
            if bold:
                etree.SubElement(f, N + "b")
            etree.SubElement(f, N + "color", rgb=font_color)
        fl = etree.SubElement(d, N + "fill")
        pf = etree.SubElement(fl, N + "patternFill", patternType=fill_pattern)
        if fill_pattern == "solid":
            etree.SubElement(pf, N + "bgColor", rgb=fg)
        else:
            etree.SubElement(pf, N + "fgColor", rgb=fg)
            etree.SubElement(pf, N + "bgColor", rgb=bg)
        self.dxfs.set("count", str(len(self.dxfs)))
        return len(self.dxfs) - 1

    def xml(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------------ sheets
ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
         "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
         "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
         "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties",
         "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
         "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


class Sheet:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(N + "sheetData")
        self.rows = {int(r.get("r")): r for r in self.data.findall(N + "row")}
        self.cells = {}
        for r in self.rows.values():
            for c in r.findall(N + "c"):
                self.cells[c.get("r")] = c

    def cell(self, r, c):
        ref = f"{L(c)}{r}"
        if ref in self.cells:
            return self.cells[ref]
        row = self.rows.get(r)
        if row is None:
            row = etree.Element(N + "row", r=str(r))
            later = [k for k in self.rows if k > r]
            if later:
                self.rows[min(later)].addprevious(row)
            else:
                self.data.append(row)
            self.rows[r] = row
        el = etree.Element(N + "c", r=ref)
        for other in row.findall(N + "c"):
            if col_idx(other.get("r"))[0] > c:
                other.addprevious(el)
                break
        else:
            row.append(el)
        self.cells[ref] = el
        return el

    def insert(self, el):
        name = etree.QName(el).localname
        pos = ORDER.index(name)
        for i, child in enumerate(self.root):
            cname = etree.QName(child).localname
            if cname in ORDER and ORDER.index(cname) > pos:
                child.addprevious(el)
                return
        self.root.append(el)

    def add_cf(self, sqref, formula, dxf, priority, stop=False):
        cf = etree.Element(N + "conditionalFormatting", sqref=sqref)
        rule = etree.SubElement(cf, N + "cfRule", type="expression", dxfId=str(dxf), priority=str(priority))
        if stop:
            rule.set("stopIfTrue", "1")
        etree.SubElement(rule, N + "formula").text = formula
        existing = self.root.findall(N + "conditionalFormatting")
        if existing:
            existing[-1].addnext(cf)
        else:
            self.insert(cf)

    def add_dv(self, sqref, formula, title, error):
        dvs = self.root.find(N + "dataValidations")
        if dvs is None:
            dvs = etree.Element(N + "dataValidations", count="0")
            self.insert(dvs)
        dv = etree.SubElement(dvs, N + "dataValidation", type="custom", allowBlank="1", showErrorMessage="1",
                              errorTitle=title, error=error, sqref=sqref)
        etree.SubElement(dv, N + "formula1").text = formula
        dvs.set("count", str(len(dvs)))

    def set_text(self, r, c, text):
        el = self.cell(r, c)
        for ch in list(el):
            el.remove(ch)
        el.set("t", "inlineStr")
        is_ = etree.SubElement(el, N + "is")
        etree.SubElement(is_, N + "t").text = text

    def xml(self):
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------------ layout readers
def label_grid(ws, cols, first, last):
    top = {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                top[(r, c)] = (m.min_row, m.min_col)

    def v(r, c):
        rr, cc = top.get((r, c), (r, c))
        x = ws.cell(rr, cc).value
        return re.sub(r"\s+", " ", str(x)).strip() if x is not None else ""
    return {r: [v(r, c) for c in cols] for r in range(first, last + 1)}


SHEETS = {
    # name: label cols (overlay, location, resource, origin, rank), data first/last col, rows, additional rows,
    #       days per risk, totals cols
    "Match Day Resource": dict(labels=[3, 4, 5, 6, 7], c0=8, c1=55, rows=(8, 344), add=(347, 371), days=4,
                               tot=(56, 67), head_rows=(6,)),
    "Wider Tournament Resource": dict(labels=[2, 3, 4, 5, 6], c0=7, c1=18, rows=(6, 90), add=(93, 102), days=1,
                                      tot=None, head_rows=(3,)),
    "Pan London": dict(labels=[1, 2, 3, 4, 5], c0=6, c1=17, rows=(6, 34), add=(37, 46), days=1, tot=None,
                       head_rows=(3,)),
}
MD_LOCS = {"In Stadium", "Stadium Vicinity", "Town Centre"}


def classify(name, spec, ws):
    """row -> dict(kind, hatch set of (day, measure))"""
    grid = label_grid(ws, spec["labels"], spec["rows"][0], spec["rows"][1])
    out = {}
    for r, (overlay, loc, res, origin, rank) in grid.items():
        if not res or rank in ("Rank", "") and not res:
            continue
        if len({overlay, loc, res, origin, rank} - {""}) == 1:      # section heading row
            continue
        if rank == "Rank":
            continue
        rl, ol = res.lower(), origin.lower()
        hatch = set()
        days = range(spec["days"])
        if spec["days"] == 4:
            nonmd = bool(re.search(r"non-? ?match ?day", rl))
            mdonly = (not nonmd) and ("(match day" in rl or (r <= 48 and loc in MD_LOCS))
            for d in days:
                if (mdonly and d != 2) or (nonmd and d == 2):
                    hatch |= {(d, m) for m in range(3)}
        unit = rank in ("Level 2 PSU", "Level 3 BDU")
        ma = ol.startswith("mutual aid")
        in_stadium = loc == "In Stadium"
        insp = rank in ("Inspector", "Detective Inspector")
        for d in days:
            if unit:
                hatch |= {(d, 1), (d, 2)}
            elif in_stadium:
                hatch.add((d, 2))
            elif ma:
                hatch.add((d, 1))
            elif insp:
                hatch.add((d, 2))
        out[r] = dict(res=res, origin=origin, rank=rank, loc=loc, hatch=hatch)
    return out


def process_sheet(name, spec, ws, sheet, st, dxf):
    c0, c1, days = spec["c0"], spec["c1"], spec["days"]
    per_risk = days * 3
    rows = classify(name, spec, ws)
    hatch_fill = st.add_fill(*HATCH)
    tint_fill = {}

    def tint(kind, m):
        key = (kind, m)
        if key not in tint_fill:
            tint_fill[key] = st.add_fill("solid", (HIGH_T if kind == "H" else SH_T)[m])
        return tint_fill[key]

    risk_of = lambda c: (c - c0) // per_risk  # noqa: E731
    pos_of = lambda c: (((c - c0) % per_risk) // 3, (c - c0) % 3)  # noqa: E731
    # 1. recolour High / Super High (all rows from the header down)
    max_row = max(sheet.rows) if sheet.rows else 0
    color_cols = [c for c in range(c0, c1 + 1) if risk_of(c) in (2, 3)]
    if spec["tot"]:
        color_cols += [c for c in range(spec["tot"][0], spec["tot"][1] + 1) if (c - spec["tot"][0]) // 3 in (2, 3)]
    for r in range(spec["head_rows"][0] - 3, max_row + 1):
        for c in color_cols:
            ref = f"{L(c)}{r}"
            el = sheet.cells.get(ref)
            if el is None or el.get("s") is None:
                continue
            s = int(el.get("s"))
            info = st.fill_rgb(s)
            if not info or info[0] != "solid":
                continue
            kind = "H" if (risk_of(c) == 2 if c <= c1 else (c - spec["tot"][0]) // 3 == 2) else "S"
            if info[1] in PINK or (name != "Match Day Resource" and info[1] == "FFEBF1FA" and
                                   r not in spec["head_rows"]):
                m = PINK.get(info[1], (c - c0) % 3)
                el.set("s", str(st.variant(s, fill_id=tint(kind, m))))
            elif r in spec["head_rows"] and info[1] in HEAD3:
                el.set("s", str(st.variant(s, fill_id=st.add_fill("solid", HEAD3[info[1]]))))
    # Low / Medium header tints on WT / Pan London row 3
    if name != "Match Day Resource":
        for c in range(c0, c1 + 1):
            el = sheet.cells.get(f"{L(c)}{spec['head_rows'][0]}")
            if el is not None and el.get("s"):
                info = st.fill_rgb(int(el.get("s")))
                if info and info[1] in HEAD3 and risk_of(c) in (0, 1):
                    el.set("s", str(st.variant(int(el.get("s")), fill_id=st.add_fill("solid", HEAD3[info[1]]))))
    # 2. hatch + lock / unlock
    hatched, open_cells = set(), {1: set(), 2: set()}
    for r, info in rows.items():
        for c in range(c0, c1 + 1):
            d, m = pos_of(c)
            el = sheet.cell(r, c)
            s = int(el.get("s", 0))
            if (d, m) in info["hatch"]:
                el.set("s", str(st.variant(s, fill_id=hatch_fill, locked=True)))
                hatched.add((r, c))
            else:
                el.set("s", str(st.variant(s, locked=False)))
                if m in (1, 2):
                    open_cells[m].add((r, c))
    # additional resources rows: unlock labels + data, dynamic rules
    a0, a1 = spec["add"]
    lab_res, lab_origin, lab_rank = spec["labels"][2], spec["labels"][3], spec["labels"][4]
    add_open = {1: set(), 2: set()}
    for r in range(a0, a1 + 1):
        for c in [lab_res, lab_origin, lab_rank] + list(range(c0, c1 + 1)):
            el = sheet.cell(r, c)
            el.set("s", str(st.variant(int(el.get("s", 0)), locked=False)))
            if c >= c0 and pos_of(c)[1] in (1, 2):
                add_open[pos_of(c)[1]].add((r, c))
                open_cells[pos_of(c)[1]].add((r, c))
    o, k = L(lab_origin), L(lab_rank)
    opp = rects(add_open[1])
    chg = rects(add_open[2])
    first_opp, first_chg = opp[0].split(":")[0], chg[0].split(":")[0]
    ra = col_idx(first_opp)[1]
    ma_rule = f'LEFT(${o}{ra},6)="Mutual"'
    insp_rule = f'AND(LEFT(${o}{ra},5)="Force",OR(${k}{ra}="Inspector",${k}{ra}="Detective Inspector"))'
    sheet.add_cf(" ".join(opp), ma_rule, dxf["hatch"], 1, stop=True)
    sheet.add_cf(" ".join(chg), insp_rule, dxf["hatch"], 2, stop=True)
    sheet.add_dv(" ".join(opp), f"NOT({ma_rule})", "Not required",
                 "Opportunity Cost is not required for Mutual Aid lines.")
    sheet.add_dv(" ".join(chg), f"NOT({insp_rule})", "Not required",
                 "Chargeable Hours are not required for Force Supplied Inspectors.")
    # 3. block entry in hatched cells
    hr = rects(hatched)
    for i in range(0, len(hr), 400):
        sheet.add_dv(" ".join(hr[i:i + 400]), "FALSE", "Not required",
                     "This cell is not required for this resource, rank or day. Please leave it blank.")
    # 4. missing-hours check on open hours cells
    for m, prio in ((1, 3), (2, 4)):
        rr = rects(open_cells[m])
        fr, fc = col_idx(rr[0].split(":")[0])[1], col_idx(rr[0].split(":")[0])[0]
        off = f"{L(fc - m)}{fr}"
        me = f"{L(fc)}{fr}"
        sheet.add_cf(" ".join(rr), f"AND(N({off})>0,ISBLANK({me}))", dxf["missing"], prio)
    return rows, hatched


def guidance_sheet_xml(files, met):
    """Builds the guidance with openpyxl and converts it for injection (styles merged into files)."""
    buf = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output", "_guidance_tmp.xlsx")
    make_guidance.build(buf, met=met)
    gz = zipfile.ZipFile(buf)
    gfiles = {n: gz.read(n) for n in gz.namelist()}
    os.remove(buf)
    styles, off = inject.merge_styles(files["xl/styles.xml"].decode(), gfiles["xl/styles.xml"].decode())
    files["xl/styles.xml"] = styles.encode()
    sst = inject.shared_strings(gfiles.get("xl/sharedStrings.xml", b"<sst/>").decode())
    return inject.convert_sheet(gfiles["xl/worksheets/sheet1.xml"].decode(), sst, off).encode()


def remove_sheet(files, name):
    wb = files["xl/workbook.xml"].decode()
    names = re.findall(r'<sheet\b[^>]*name="([^"]+)"', wb)
    idx = names.index(name)
    part = inject.sheet_part(files, name)
    tag = re.search(rf'<sheet\b[^>]*name="{re.escape(name)}"[^>]*/>', wb).group(0)
    rid = re.search(r'r:id="([^"]+)"', tag).group(1)
    wb = wb.replace(tag, "")

    def fix(m):
        i = int(m.group(1))
        return "" if i == idx else m.group(0).replace(f'localSheetId="{i}"', f'localSheetId="{i - 1 if i > idx else i}"')
    wb = re.sub(r'<definedName\b[^>]*localSheetId="(\d+)"[^>]*>[^<]*</definedName>', fix, wb)
    wb = re.sub(r'activeTab="\d+"', 'activeTab="1"', wb)
    files["xl/workbook.xml"] = wb.encode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    rels = re.sub(rf'<Relationship\b[^>]*Id="{rid}"[^>]*/>', "", rels)
    files["xl/_rels/workbook.xml.rels"] = rels.encode()
    srels = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    ct = files["[Content_Types].xml"].decode()
    for p in [part] + ([srels] if srels in files else []):
        ct = ct.replace(f'<Override PartName="/{p}" ContentType="application/vnd.openxmlformats-officedocument.'
                        f'spreadsheetml.worksheet+xml"/>', "")
    if srels in files:
        for target in re.findall(r'Target="\.\./drawings/(drawing\d+\.xml)"', files[srels].decode()):
            dp = f"xl/drawings/{target}"
            files.pop(dp, None)
            files.pop(f"xl/drawings/_rels/{target}.rels", None)
            ct = re.sub(rf'<Override PartName="/{dp}"[^>]*/>', "", ct)
        files.pop(srels)
    files.pop(part)
    files["[Content_Types].xml"] = ct.encode()
    drop_calc_chain(files)
    if "docProps/app.xml" in files:
        app = files["docProps/app.xml"].decode()
        app = app.replace(f"<vt:lpstr>{name}</vt:lpstr>", "")
        app = re.sub(r'(<vt:vector size=")(\d+)(" baseType="lpstr">)',
                     lambda m: m.group(1) + str(int(m.group(2)) - 1) + m.group(3), app, count=1)
        app = re.sub(r"(<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>)(\d+)",
                     lambda m: m.group(1) + str(int(m.group(2)) - 1), app)
        files["docProps/app.xml"] = app.encode()


def drop_calc_chain(files):
    if "xl/calcChain.xml" in files:
        files.pop("xl/calcChain.xml")
        files["[Content_Types].xml"] = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"",
                                              files["[Content_Types].xml"])
        files["xl/_rels/workbook.xml.rels"] = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"",
                                                     files["xl/_rels/workbook.xml.rels"])


def build(src, out, met):
    z = zipfile.ZipFile(src)
    files = {n: z.read(n) for n in z.namelist()}
    wbr = openpyxl.load_workbook(src)
    # guidance first (its styles are appended to styles.xml)
    files[inject.sheet_part(files, "Guidance")] = guidance_sheet_xml(files, met)
    st = Styles(files["xl/styles.xml"])
    dxf = {"hatch": st.add_dxf(HATCH[0], HATCH[1], HATCH[2]),
           "missing": st.add_dxf("solid", "FFFFC7CE", font_color="FF9C0006", bold=True)}
    summary = {}
    for name, spec in SHEETS.items():
        if name == "Pan London" and not met:
            continue
        part = inject.sheet_part(files, name)
        sheet = Sheet(files[part])
        rows, hatched = process_sheet(name, spec, wbr[name], sheet, st, dxf)
        summary[name] = (len(rows), len(hatched))
        if name == "Match Day Resource":
            for i, (lab, color) in enumerate([("Low", "FF2E7D32"), ("Medium", "FFD4A017"), ("High", "FFBF5B04"),
                                             ("Super High", "FF8B0000")]):
                c = 56 + i * 3
                sheet.set_text(4, c, lab)
                el = sheet.cell(4, c)
                el.set("s", str(st.variant(int(el.get("s", 0)), fill_id=st.add_fill("solid", color))))
        files[part] = sheet.xml()
    files["xl/styles.xml"] = st.xml()
    sst = files["xl/sharedStrings.xml"].decode()
    sst = sst.replace("requered", "required").replace("Catagories", "Categories")
    files["xl/sharedStrings.xml"] = sst.encode()
    if not met:
        remove_sheet(files, "Pan London")
    wb = files["xl/workbook.xml"].decode()
    if "fullCalcOnLoad" not in wb:
        wb = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wb)
    files["xl/workbook.xml"] = wb.encode()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n in z.namelist():
            if n in files:
                zo.writestr(z.getinfo(n), files[n])
    return summary


if __name__ == "__main__":
    src, out_dir = sys.argv[1], sys.argv[2]
    os.makedirs(out_dir, exist_ok=True)
    for met, fname in ((False, "Resource_Collection_Template_V2_General.xlsx"),
                       (True, "Resource_Collection_Template_V2_MPS_Pan_London.xlsx")):
        print(fname, build(src, os.path.join(out_dir, fname), met))
