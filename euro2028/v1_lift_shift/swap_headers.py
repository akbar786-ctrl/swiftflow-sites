"""
Swaps header rows 3 and 4 on 'Match Day Resource' so the RISK LEVEL is on top:

  before  row 3: Match Day -2 | Match Day -1 | Match Day | Match Day +1
          row 4: Low-Number required, Low-Normal ToD, ... (per day)
  after   row 3: Low | Medium | High
          row 4: Match Day -2 - Number required, ... (per risk level)

Columns D:AM are reordered cell by cell (values, styles, hatching move with
them) directly in the sheet XML, so nothing else in the template changes.
The Automatic Total Figures (AN:AV) are rewritten as direct sums because the
original SUMIF matched on the row 4 labels.
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


def transform(xml):
    def row(m):
        rnum = int(m.group(1))
        head, body, tail = m.group(0)[:m.start(2) - m.start(0)], m.group(2), "</row>"
        cells = CELL.findall(body)
        out = []
        for c in cells:
            ref = re.search(r'r="([A-Z]+)(\d+)"', c)
            col = CI(ref.group(1))
            if rnum >= 5 and col in PERM:                       # data rows: move the cell
                c = c.replace(f'r="{ref.group(1)}{rnum}"', f'r="{L(PERM[col])}{rnum}"', 1)
                col = PERM[col]
            elif rnum == 3 and 4 <= col <= 39:                  # row 3: risk level banners
                off = col - 4
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
