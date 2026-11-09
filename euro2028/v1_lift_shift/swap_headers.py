"""
Swaps header rows 3 and 4 on 'Match Day Resource' so the RISK LEVEL is on top:

  before  row 3: Match Day -2 | Match Day -1 | Match Day | Match Day +1
          row 4: Low-Number required, Low-Normal ToD, ... (per day)
  after   row 3: Low | Medium | High
          row 4: Match Day -2 - Number required, ... (per risk level)

Columns D:AM are reordered cell by cell (values, styles, hatching move with
them) directly in the sheet XML, so nothing else in the template changes.
The Automatic Total Figures are rewritten as direct sums because the
original SUMIF matched on the row 4 labels. A Super High block (AN:AY) is then
inserted after High, styled like High, and the totals move to AZ:BK.
"""
import re

from openpyxl.utils import column_index_from_string as CI
from openpyxl.utils import get_column_letter as L

DAYS = ["Match Day -2", "Match Day -1", "Match Day", "Match Day +1"]
RISKS = ["Low", "Medium", "High"]
MEASURES = ["Number required", "Normal ToD", "Extra Hours"]


def old_col(d, s, m):          # day-first (original)
    return 4 + d * 9 + s * 3 + m


def new_col(d, s, m):          # risk-first (swapped)
    return 4 + s * 12 + d * 3 + m


PERM = {old_col(d, s, m): new_col(d, s, m) for d in range(4) for s in range(3) for m in range(3)}


def _esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _set_text(cell_xml, text):
    """Return cell XML with an inline string value, keeping its style."""
    attrs = re.match(r"<c\b((?:[^>/]|/(?!>))*)", cell_xml).group(1)
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    return f'<c{attrs} t="inlineStr"><is><t xml:space="preserve">{_esc(text)}</t></is></c>'


def _clear(cell_xml):
    attrs = re.match(r"<c\b((?:[^>/]|/(?!>))*)", cell_xml).group(1)
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    return f"<c{attrs}/>"


CELL = re.compile(r"<c\b(?:[^>/]|/(?!>))*(?:/>|>.*?</c>)", re.S)


def swap(xml):
    def row(m):
        rnum = int(m.group(1))
        head, body, tail = m.group(0)[:m.start(2) - m.start(0)], m.group(2), "</row>"
        cells = CELL.findall(body)
        d3_style = next((re.search(r'\bs="(\d+)"', c) for c in cells if re.search(r'r="D3"', c)), None)
        out = []
        for c in cells:
            ref = re.search(r'r="([A-Z]+)(\d+)"', c)
            col = CI(ref.group(1))
            if rnum >= 4 and col in PERM:                       # row 4 and data rows: move the cell
                c = c.replace(f'r="{ref.group(1)}{rnum}"', f'r="{L(PERM[col])}{rnum}"', 1)  # (colours follow
                col = PERM[col]                                                            #  their risk level)
            if rnum == 3 and 4 <= col <= 39:                    # row 3: risk level banners, styled like D3
                off = col - 4
                if d3_style:
                    c = re.sub(r'\bs="\d+"', f's="{d3_style.group(1)}"', c, count=1) if ' s="' in c else \
                        c.replace("<c ", f'<c s="{d3_style.group(1)}" ', 1)
                c = _set_text(c, RISKS[off // 12]) if off % 12 == 0 else _clear(c)
            elif rnum == 4 and 4 <= col <= 39:                  # row 4: day - measure
                off = col - 4
                s, rest = divmod(off, 12)
                d, mm = divmod(rest, 3)
                c = _set_text(c, f"{DAYS[d]} - {MEASURES[mm]}")
            if rnum >= 5 and 40 <= col <= 48 and "<f" in c:    # AN:AV totals -> direct sums
                s, mm = divmod(col - 40, 3)
                refs = ",".join(f"{L(new_col(d, s, mm))}{rnum}" for d in range(4))
                c = re.sub(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", f"<f>SUM({refs})</f>", c, count=1, flags=re.S)
            out.append((col, c))
        out.sort(key=lambda x: x[0])
        return head + "".join(c for _, c in out) + tail

    # (?:[^>/]|/(?!>))* stops a self-closing <row .../> from swallowing the next row
    xml = re.sub(r'<row r="(\d+)"(?:[^>/]|/(?!>))*>(.*?)</row>', row, xml, flags=re.S)
    # header merges on row 3: per-day -> per-risk
    for old in ("D3:L3", "M3:U3", "V3:AD3", "AE3:AM3"):
        xml = xml.replace(f'<mergeCell ref="{old}"/>', "", 1)
    new = "".join(f'<mergeCell ref="{L(4 + s * 12)}3:{L(15 + s * 12)}3"/>' for s in range(3))
    xml = re.sub(r'(<mergeCells count=")(\d+)(">)', lambda m: f"{m.group(1)}{int(m.group(2)) - 1}{m.group(3)}{new}", xml, count=1)
    return xml


# ------------------------------------------------------------------ Super High
# After the swap, a fourth risk level is inserted after High:
#   AN:AY  Super High (Match Day -2 / -1 / Match Day / +1 x Number | Normal ToD | Extra Hours)
#   AZ:BK  Automatic Total Figures (Low, Medium, High, Super High x 3 measures)
# Old AN:AV (totals) move 12 to the right; everything from AW onwards moves 15.
RISKS4 = RISKS + ["Super High"]
SH_FIRST, SH_LAST = 40, 51            # AN:AY
TOT_FIRST, TOT_LAST = 52, 63          # AZ:BK


def shift(col):
    if 40 <= col <= 48:
        return col + 12
    if col >= 49:
        return col + 15
    return col


def map_range(ref, extend=True):
    """Shift a cell/range reference; ranges ending at AM or AV are widened to AY / BK."""
    def one(part):
        m = re.fullmatch(r"(\$?)([A-Z]+)(\$?)(\d+)", part)
        return m.group(1), CI(m.group(2)), m.group(3), m.group(4)
    out = []
    for piece in ref.split():
        ends = piece.split(":")
        a = one(ends[0])
        a_col = shift(a[1]) if a[1] < 16384 - 15 else a[1]
        if len(ends) == 1:
            out.append(f"{a[0]}{L(a_col)}{a[2]}{a[3]}")
            continue
        b = one(ends[1])
        if b[1] >= 16384 - 15:
            b_col = b[1]
        elif extend and b[1] == 39 and a[1] <= 39:
            b_col = SH_LAST
        elif extend and b[1] == 48 and a[1] <= 48:
            b_col = TOT_LAST
        else:
            b_col = shift(b[1])
        out.append(f"{a[0]}{L(a_col)}{a[2]}{a[3]}:{b[0]}{L(b_col)}{b[2]}{b[3]}")
    return " ".join(out)


def add_super_high(xml):
    def row(m):
        rnum = int(m.group(1))
        head = re.sub(r'\s+spans="[^"]*"', "", m.group(0)[:m.start(2) - m.start(0)])
        cells = {}
        for c in CELL.findall(m.group(2)):
            ref = re.search(r'r="([A-Z]+)(\d+)"', c)
            col = CI(ref.group(1))
            new = shift(col)
            if new != col:
                c = c.replace(f'r="{ref.group(1)}{rnum}"', f'r="{L(new)}{rnum}"', 1)
            cells[new] = c

        def copy(src, dst):
            if src in cells:
                c = cells[src].replace(f'r="{L(src)}{rnum}"', f'r="{L(dst)}{rnum}"', 1)
                cells[dst] = _clear(c)
        if rnum >= 3:                                          # (rows 1-2 are the banner)
            for i in range(12):                                # Super High inputs styled like High
                copy(28 + i, SH_FIRST + i)
            for i in range(3):                                 # Super High totals styled like High totals
                copy(TOT_FIRST + 6 + i, TOT_FIRST + 9 + i)
        if rnum == 3 and SH_FIRST in cells:
            cells[SH_FIRST] = _set_text(cells[SH_FIRST], "Super High")
        if rnum == 4:
            for i in range(12):
                d, mm = divmod(i, 3)
                if SH_FIRST + i in cells:
                    cells[SH_FIRST + i] = _set_text(cells[SH_FIRST + i], f"{DAYS[d]} - {MEASURES[mm]}")
                s, mm = divmod(i, 3)
                if TOT_FIRST + i in cells:
                    cells[TOT_FIRST + i] = _set_text(cells[TOT_FIRST + i], f"{RISKS4[s]} - {MEASURES[mm]}")
        if rnum >= 5 and TOT_FIRST in cells and "<f" in cells[TOT_FIRST]:
            for i in range(12):                                # totals: direct sums of the 4 day columns
                s, mm = divmod(i, 3)
                refs = ",".join(f"{L(new_col(d, s, mm))}{rnum}" for d in range(4))
                attrs = re.match(r"<c\b((?:[^>/]|/(?!>))*)", cells.get(TOT_FIRST + i, cells[TOT_FIRST])).group(1)
                attrs = re.sub(r'\s+t="[^"]*"', "", attrs).replace(f'r="{L(TOT_FIRST)}{rnum}"', f'r="{L(TOT_FIRST + i)}{rnum}"')
                cells[TOT_FIRST + i] = f"<c{attrs}><f>SUM({refs})</f></c>"
        return head + "".join(cells[k] for k in sorted(cells)) + "</row>"

    xml = re.sub(r'<row r="(\d+)"(?:[^>/]|/(?!>))*>(.*?)</row>', row, xml, flags=re.S)
    # column widths: D as before, E:BK input/total width, then the narrow columns
    xml = re.sub(r"<cols>.*?</cols>", lambda m: re.sub(
        r'<col min="5" max="48"([^>]*)/><col min="49" max="79"',
        rf'<col min="5" max="{TOT_LAST}"\1/><col min="{TOT_LAST + 1}" max="{79 + 15}"', m.group(0)), xml, count=1, flags=re.S)
    xml = re.sub(r'<dimension ref="A1:([A-Z]+)(\d+)"/>', lambda m: f'<dimension ref="A1:{L(shift(CI(m.group(1))))}{m.group(2)}"/>', xml)
    # merges: shift only; totals banner widened; Super High banner added
    def merge(m):
        ref = m.group(1)
        if ref == "AN3:AV3":
            return f'<mergeCell ref="{L(TOT_FIRST)}3:{L(TOT_LAST)}3"/>'
        return f'<mergeCell ref="{map_range(ref, extend=False)}"/>'
    xml = re.sub(r'<mergeCell ref="([^"]+)"/>', merge, xml)
    xml = re.sub(r'(<mergeCells count=")(\d+)(">)',
                 lambda m: f'{m.group(1)}{int(m.group(2)) + 1}{m.group(3)}<mergeCell ref="{L(SH_FIRST)}3:{L(SH_LAST)}3"/>', xml, count=1)
    # conditional formatting, validations, ignored errors
    xml = re.sub(r'(<(?:conditionalFormatting|dataValidation|ignoredError)\b[^>]*\bsqref=")([^"]+)"',
                 lambda m: m.group(1) + map_range(m.group(2)) + '"', xml)
    xml = re.sub(r"<xm:sqref>([^<]+)</xm:sqref>", lambda m: f"<xm:sqref>{map_range(m.group(1))}</xm:sqref>", xml)
    # filter: same range, buttons hidden on D:BJ as the headers are merged banners
    xml = re.sub(r'<autoFilter ref="([^"]+)"([^>]*)>.*?</autoFilter>',
                 lambda m: f'<autoFilter ref="{map_range(m.group(1))}"{m.group(2)}>'
                 + "".join(f'<filterColumn colId="{i}" showButton="0"/>' for i in range(3, TOT_LAST - 1))
                 + "</autoFilter>", xml, flags=re.S)
    return xml


def transform(xml):
    return add_super_high(swap(xml))


def workbook(xml):
    """Keep the Match Day filter range in step with the wider sheet."""
    return xml.replace("'Match Day Resource'!$A$3:$AV$328", f"'Match Day Resource'!$A$3:${L(TOT_LAST)}$328")
