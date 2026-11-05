"""
Row layout of the NEW template v1 (BLANK_Euro_2028_Resource_NEW.xlsx) and the
legacy calculation logic that applies to each row.

Columns are identical to the original collection survey:
  Match Day Resource         D:L MD-2, M:U MD-1, V:AD Match Day, AE:AM MD+1,
                             each Low | Medium | High x Number | Normal ToD | Extra Hours
  Wider Tournament Resource  D:O Low | Medium | High | Super High x Number | Normal ToD | Extra Hours

verify(path) checks every row used here against the template's own labels.
"""
import openpyxl

MD_SHEET, WT_SHEET = "Match Day Resource", "Wider Tournament Resource"
SCEN = ["Low", "Medium", "High"]
WT_SCEN = ["Low", "Medium", "High", "Super High"]
# Match Day Resource columns after the header swap (risk level on top, then day, then measure):
#   D:O Low, P:AA Medium, AB:AM High; within each, MD-2 | MD-1 | MD | MD+1 x Number | ToD | Extra Hours
DAY_IDX = {"MD-2": 0, "MD-1": 1, "MD": 2, "MD+1": 3}


def md_col(day, s, m=0):
    """Column number on Match Day Resource for relative day, risk index s (0 Low..2 High), measure m (0 Number, 1 ToD, 2 Extra)."""
    return 4 + s * 12 + DAY_IDX[day] * 3 + m

# ---------------------------------------------------------------- Match Day
# PSU / serial configurations: first row of each block (Inspector, Sergeant, Constable)
MD_CONFIG = {7: "Level 2", 10: "Level 3", 13: "Level 2", 16: "Level 3",
             19: "Level 2", 22: "Level 3", 25: "Level 2", 28: "Level 3"}
# In-Stadium PSU configurations: legacy charged their Inspectors at the MA Inspector rate
MD_IN_STADIUM_CONFIG = {25, 28}
# Public order counts: row -> (supply, group, configuration first row)
#   group NONMD : MD-1 > MD-2 > MD+1  (Match Day not used)
#   group MDONLY: Match Day only
MD_COUNTS = {
    33: ("FS", "NONMD", 7), 34: ("FS", "NONMD", 10), 35: ("MA", "NONMD", None), 36: ("MA", "NONMD", None),
    37: ("FS", "MDONLY", 13), 38: ("FS", "MDONLY", 16), 39: ("MA", "MDONLY", None), 40: ("MA", "MDONLY", None),
    41: ("FS", "MDONLY", 19), 42: ("FS", "MDONLY", 22), 43: ("MA", "MDONLY", None), 44: ("MA", "MDONLY", None),
    45: ("FS", "MDONLY", 25), 46: ("FS", "MDONLY", 28), 47: ("MA", "MDONLY", None), 48: ("MA", "MDONLY", None),
}
# Officer rows (group STD: MD > MD-1 > MD-2 > MD+1)
MD_OFFICER_BLOCKS = [(50, 53), (55, 72), (74, 85), (88, 91), (94, 117), (120, 127), (130, 141), (144, 175),
                     (177, 190), (193, 204), (207, 224), (227, 256), (259, 270), (273, 288), (291, 297),
                     (300, 302), (305, 306), (309, 314), (317, 322), (325, 328), (331, 339)]
MD_ROLE_ROWS = set(range(74, 86)) | set(range(177, 191)) | set(range(279, 289))     # Role Profile drop-down
MD_RANK_DROPDOWN_ROWS = set(range(88, 92)) | set(range(94, 118)) | set(range(130, 142))
MD_ADDITIONAL_ROWS = set(range(331, 340))
MD_SECTIONS = [6, 32, 49, 54, 73, 87, 93, 119, 129, 143, 176, 192, 206, 226, 258, 272, 290, 299, 304, 308,
               316, 324, 330]

# ------------------------------------------------------- Wider Tournament
WT_CONFIG = {6: "Level 2", 9: "Level 3", 12: "Level 2", 15: "Level 3"}
WT_COUNTS = {20: ("FS", 6), 21: ("FS", 9), 22: ("MA", None), 23: ("MA", None),
             24: ("FS", 12), 25: ("FS", 15), 26: ("MA", None), 27: ("MA", None)}
WT_OFFICER_BLOCKS = [(28, 37), (39, 56), (59, 70), (73, 76), (79, 90), (93, 101)]
WT_ROLE_ROWS = set(range(39, 57))
WT_ADDITIONAL_ROWS = set(range(93, 102))
WT_SECTIONS = [5, 19, 38, 58, 72, 78, 92]

RANK_NORMALISE = {"sergeants": "Sergeant", "constables": "Constable", "supt": "Superintendent",
                  "detecive constable": "Detective Constable"}


def norm_rank(text):
    t = (text or "").replace("\xa0", " ").strip()
    return RANK_NORMALISE.get(t.lower(), t)


def md_officer_rows():
    return [r for a, b in MD_OFFICER_BLOCKS for r in range(a, b + 1)]


def wt_officer_rows():
    return [r for a, b in WT_OFFICER_BLOCKS for r in range(a, b + 1)]


class Template:
    """Reads labels from the v1 template (read-only) so formulas can carry them."""

    def __init__(self, path):
        wb = openpyxl.load_workbook(path)
        self.md, self.wt = wb[MD_SHEET], wb[WT_SHEET]
        self.block_label = {MD_SHEET: self._blocks(self.md), WT_SHEET: self._blocks(self.wt)}

    @staticmethod
    def _blocks(ws):
        """row -> resource label (column B, carried down merged blocks)."""
        out = {}
        for m in ws.merged_cells.ranges:
            if m.min_col == 2 and m.max_col == 2:
                for r in range(m.min_row, m.max_row + 1):
                    out[r] = ws.cell(m.min_row, 2).value
        for r in range(1, ws.max_row + 1):
            if r not in out and ws.cell(r, 2).value:
                out[r] = ws.cell(r, 2).value
        return {r: (v or "").replace("\xa0", " ").strip() for r, v in out.items()}

    def label(self, sheet, r):
        return self.block_label[sheet].get(r, "")

    def supply(self, sheet, r):
        lab = self.label(sheet, r).lower()
        if lab.startswith("mutual aid"):
            return "MA"
        if lab.startswith("force supplied"):
            return "FS"
        return None

    def rank(self, sheet, r):
        ws = self.md if sheet == MD_SHEET else self.wt
        return norm_rank(ws.cell(r, 3).value)

    def section_title(self, sheet, r):
        ws = self.md if sheet == MD_SHEET else self.wt
        a = (ws.cell(r, 1).value or "").strip()
        if not a or "drop down" in a.lower():
            a = (ws.cell(r, 2).value or a).strip()
        return a


def verify(path):
    t = Template(path)
    problems = []
    for r in MD_CONFIG:
        for i, rk in enumerate(["Inspector", "Sergeant", "Constable"]):
            if t.rank(MD_SHEET, r + i) != rk:
                problems.append(f"MD config row {r + i}: expected {rk}, found {t.rank(MD_SHEET, r + i)}")
        if MD_CONFIG[r] not in t.label(MD_SHEET, r):
            problems.append(f"MD config row {r}: '{t.label(MD_SHEET, r)}'")
    for r, (sup, grp, cfg) in MD_COUNTS.items():
        if t.supply(MD_SHEET, r) != sup:
            problems.append(f"MD count row {r}: supply {t.supply(MD_SHEET, r)} != {sup}")
        if ("Non-Matchday" in t.label(MD_SHEET, r)) != (grp == "NONMD"):
            problems.append(f"MD count row {r}: group {grp} vs '{t.label(MD_SHEET, r)}'")
    for r in md_officer_rows():
        if r not in MD_ADDITIONAL_ROWS and t.supply(MD_SHEET, r) is None:
            problems.append(f"MD officer row {r}: no Force Supplied / Mutual Aid label")
        if r not in MD_ROLE_ROWS | MD_RANK_DROPDOWN_ROWS | MD_ADDITIONAL_ROWS and not t.rank(MD_SHEET, r):
            problems.append(f"MD officer row {r}: no rank")
    for r in WT_CONFIG:
        for i, rk in enumerate(["Inspector", "Sergeant", "Constable"]):
            if t.rank(WT_SHEET, r + i) != rk:
                problems.append(f"WT config row {r + i}: expected {rk}")
    for r, (sup, cfg) in WT_COUNTS.items():
        if t.supply(WT_SHEET, r) != sup:
            problems.append(f"WT count row {r}: supply {t.supply(WT_SHEET, r)} != {sup}")
    for r in wt_officer_rows():
        if r not in WT_ADDITIONAL_ROWS and t.supply(WT_SHEET, r) is None:
            problems.append(f"WT officer row {r}: no supply label")
        if r not in WT_ROLE_ROWS | WT_ADDITIONAL_ROWS and not t.rank(WT_SHEET, r):
            problems.append(f"WT officer row {r}: no rank")
    return problems
