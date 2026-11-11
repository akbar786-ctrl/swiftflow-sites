"""
EURO 2028 collection template - "one row per resource / origin / rank" layout (v3),
following the user's redesign of Match Day Resource (Public Order and Evidence Gathering),
extended to every resource and to Wider Tournament Resource.

Layout (both sheets)
  B  Section (hidden helper column)      C Overlay   D Deployment Location   E Resource
  F  Resource Origin                     G Rank      H.. numbers
  Match Day: H:BC = Low / Medium / High / Super High x MD-2 / MD-1 / MD / MD+1 x
             Officers / Normal ToD / Extra Hours;  BD:BO automatic totals per risk level
  Wider Tournament: H:S = Low / Medium / High / Super High x Officers / Normal ToD / Extra Hours
  Row 3 risk level, row 4 day, row 5 measure, row 6 HIDDEN single header row for Power BI
  ("Low | MD-2 | Officers" ...), data from row 8 (section bars between groups).

Usage: python3 make_template.py [out.xlsx]
"""
import io
import os
import sys
import zipfile

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
BLANK = os.path.join(HERE, "..", "v1_lift_shift", "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "output", "EURO_2028_Resource_Collection_Template_v3.xlsx")

FS, MA = "Force Supplied", "Mutual Aid Supplied"
NA = "N/A"
RISKS = ["Low", "Medium", "High", "Super High"]
DAYS = [("MD-2", "Match Day -2"), ("MD-1", "Match Day -1"), ("MD", "Match Day"), ("MD+1", "Match Day +1")]
MEASURES = [("Officers", "Total Officers required", "Officers"), ("Normal ToD", "Normal Tour of Duty", "Hours"),
            ("Extra Hours", "Extra Hours / Over Time", "Hours")]
# banner, day banner, cell tint, banner text colour
RISK_COL = {"Low": ("548235", "70AD47", "E2EFDA", "FFFFFF"), "Medium": ("FFC000", "FFD966", "FFF2CC", "000000"),
            "High": ("C55A11", "F4B183", "FBE5D6", "FFFFFF"), "Super High": ("C00000", "FF7C80", "FFE1E1", "FFFFFF")}
TOTAL_COL = ("7030A0", "B4A7D6", "EDE7F6", "FFFFFF")
NAVY, SECTION = "1F3864", "1F4E79"
FONT = "Aptos Narrow"
ISC = ["Inspector", "Sergeant", "Constable"]
SC = ["Sergeant", "Constable"]
RANKS = ["Assistant Chief Constable", "Chief Superintendent", "Superintendent", "Chief Inspector", "Inspector",
         "Sergeant", "Constable"]

# role -> rank (from the national team's list)
PO_ROLES = [("Forward Intelligence Officer", "Constable"), ("Police Liaison Team Officer", "Constable"),
            ("Medic", "Constable"), ("Bronze Commander", "Inspector"), ("Safety Advisor", "Constable"),
            ("Loggist", "Constable")]
CBRN_ROLES = [("Enhanced Responder", "Constable"), ("Operational Commander (Bronze)", "Inspector"),
              ("Search", "Constable"), ("Tactical Advisor", "Constable"), ("Tactical Commander (Silver)", "Chief Inspector")]
ARMED_ROLES = [("Firearms Tactical Advisor", "Constable"), ("Specialist Firearms Officer", "Constable"),
               ("Specialist Strategic Firearms Commander", "Assistant Chief Constable"),
               ("Specialist Tactical Firearms Commander", "Chief Inspector"),
               ("Strategic Firearms Commander", "Chief Superintendent"), ("Armed Response Vehicle (ARV)", "Constable"),
               ("CTSFO", "Constable")]


def R(section, resource, origin, rank, overlay=NA, location=NA, free=False):
    return dict(section=section, overlay=overlay, location=location, resource=resource, origin=origin, rank=rank,
                free=free)


def by_rank(section, resource, ranks, origins=(FS, MA), **kw):
    return [R(section, resource, o, rk, **kw) for o in origins for rk in ranks]


def roles(section, pairs, origins=(FS, MA), **kw):
    return [R(section, role, o, rk, **kw) for o in origins for role, rk in pairs]


def additional(section, n=9):
    return [R(section, "", "", "", free=True) for _ in range(n)]


# ------------------------------------------------------------------ Match Day rows
PO = "Public Order"
MD_LOCS = ["In Stadium", "Stadium Vicinity", "Any other Location"]


def md_public_order():
    out = []
    for o in (FS, MA):
        for loc in MD_LOCS:
            out += [R(PO, "Level 2 PSU Officer - PC / SGT", o, rk, location=loc) for rk in ISC]
            out += [R(PO, "Public Order Protective Carrier Driver", o, rk, location=loc) for rk in SC]
    for o in (FS, MA):
        for loc in MD_LOCS:
            out += [R(PO, "Level 3 Public Order Public Safety Officer", o, rk, location=loc) for rk in ISC]
    return out


def md_rows():
    groups = [
        md_public_order(),
        by_rank("Evidence Gathering Team Officer", "Evidence Gathering Team Officer", SC),
        by_rank("Protest Removal", "Working at Height Complex Protest Removal", ISC)
        + by_rank("Protest Removal", "Non-Complex Protest Removal", ISC)
        + by_rank("Protest Removal", "Debonding Protest Removal", ISC),
        roles("Additional Public Order / Public Safety Resources", PO_ROLES),
        by_rank("Mounted", "Mounted Commander", ISC) + by_rank("Mounted", "Mounted Officer", ISC),
        by_rank("Dogs", "Dogs GDP", ISC) + by_rank("Dogs", "Dogs Public Order Serial Commander", ISC)
        + by_rank("Dogs", "Dogs Public Order", ISC) + by_rank("Dogs", "Dogs Explosives", ISC),
        by_rank("Football Policing", "Football Spotting DFO", SC) + by_rank("Football Policing", "Football Spotting OFO", SC),
        by_rank("Roads Policing", "Roads Policing Officer", ISC)
        + by_rank("Roads Policing", "Roads Policing Officer - Motorcyclists", ISC)
        + by_rank("Roads Policing", "Roads Policing Officer - VIP Motorcyclist Escort", ISC),
        by_rank("Armed Policing", "Operational Firearms Commander", ISC)
        + by_rank("Armed Policing", "Armed Response (ARVs)", ISC) + by_rank("Armed Policing", "CTSFOs", ISC)
        + by_rank("Armed Policing", "AFOs", ISC) + by_rank("Armed Policing", "Rifle Officer", SC)
        + by_rank("Armed Policing", "Protection Officer", SC),
        roles("Additional Armed Policing", ARMED_ROLES, overlay="Armed Response Overlay"),
        by_rank("Disruptive Effects", "Servator Disruptive Effects", SC)
        + by_rank("Disruptive Effects", "Servator Disruptive Effects Operational Commander", SC)
        + by_rank("Disruptive Effects", "Servator Disruptive Effects Tactical Commander", SC),
        by_rank("Search", "LSOs", ISC) + by_rank("Search", "PolSAs", ISC) + by_rank("Search", "CT Secco", ISC),
        sum((by_rank("Unmanned Aerial Systems (Drones)", r, ISC) for r in
             ["Counter UAS Bronze Commander", "Counter UAS Effector Operator", "Counter UAS Observers",
              "Counter UAS Tactical Advisor", "Drone"]), []),
        by_rank("Live Facial Recognition", "LFR Operational Commander", ISC)
        + by_rank("Live Facial Recognition", "LFR Operator", ISC),
        by_rank("CBRN", "General Responder", ISC) + roles("CBRN", [(f"CBRN - {n}", rk) for n, rk in CBRN_ROLES]),
        by_rank("Force Command (non Public Order)", "Command", RANKS, origins=(FS,)),
        by_rank("Control Room", "Control Room", ISC, origins=(FS,)),
        by_rank("Custody", "Custody", SC, origins=(FS,)),
        by_rank("Crime", "Crime", ["Detective Inspector", "Detective Sergeant", "Detective Constable"]),
        by_rank("Intelligence", "Intel", ISC),
        by_rank("Ancillary", "Ancillary Team", SC),
        additional("Additional Resources"),
    ]
    return groups


# ------------------------------------------------------------- Wider Tournament rows
WT_LOCS = ["Wider Community", "Fan Zone"]
PROTEST = ["Non-Complex Protest Removal", "Working at Height - Protest Removal", "De-Bonding - Protest Removal"]


def wt_rows():
    po = []
    for o in (FS, MA):
        for loc in WT_LOCS:
            po += [R(PO, "Level 2 PSU Officer - PC / SGT", o, rk, location=loc) for rk in ISC]
            po += [R(PO, "Public Order Protective Carrier Driver", o, rk, location=loc) for rk in SC]
    for o in (FS, MA):
        for loc in WT_LOCS:
            po += [R(PO, "Level 3 Public Order Public Safety Officer", o, rk, location=loc) for rk in ISC]
    po += by_rank(PO, "Operationally Deployable non-Public Order Policing", ISC)
    po += by_rank(PO, "Evidence Gatherers", SC)
    add = []
    for o in (FS, MA):
        for loc in WT_LOCS:
            add += [R("Additional Public Order / Public Safety Resources", role, o, rk, location=loc) for role, rk in PO_ROLES]
            add += [R("Additional Public Order / Public Safety Resources", p, o, rk, location=loc) for p in PROTEST for rk in ISC]
    fz = dict(location="Fan Zone")
    return [
        po, add,
        by_rank("Fan Zone - Disruptive Effects", "Disruptive Effects Officer", SC, **fz)
        + by_rank("Fan Zone - Disruptive Effects", "Disruptive Effects Operational Commander", SC, **fz)
        + by_rank("Fan Zone - Disruptive Effects", "Disruptive Effects Tactical Commander", SC, **fz),
        by_rank("Fan Zone Armed Policing", "Fan Zone Armed Response (ARVs)", SC, **fz),
        by_rank("Wider Tournament Policing", "Control Room", SC, origins=(FS,))
        + by_rank("Wider Tournament Policing", "Command Support", ["Chief Inspector"] + ISC)
        + by_rank("Wider Tournament Policing", "Custody", SC, origins=(FS,)),
        additional("Additional Resources"),
    ]


# --------------------------------------------------------------------- builder
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", fgColor=c)


def put(ws, r, c, v=None, bg=None, bold=False, color="000000", size=8, h="center", wrap=True, border=True):
    cell = ws.cell(r, c)
    if v is not None:
        cell.value = v
    cell.font = Font(name=FONT, size=size, bold=bold, color=color)
    if bg:
        cell.fill = fill(bg)
    cell.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    if border:
        cell.border = BOX
    return cell


def banner(ws, last_col, logos, title):
    for c in range(1, last_col + 1):
        ws.cell(1, c).fill = fill("2F5597")
        ws.cell(2, c).fill = fill("8EA9DB")
    ws.row_dimensions[1].height = 58
    ws.row_dimensions[2].height = 6
    ws["E1"] = title
    ws["E1"].font = Font(name=FONT, size=18, bold=True, color="FFFFFF")
    ws["E1"].alignment = Alignment(vertical="center")
    off = ws.cell(1, last_col - 1, "OFFICIAL")
    off.font = Font(name=FONT, size=12, bold=True, color="FFFFFF")
    off.alignment = Alignment(vertical="center", horizontal="right")
    for i, data in enumerate(logos):
        img = XLImage(io.BytesIO(data))
        ratio = 70 / img.height
        img.width, img.height = int(img.width * ratio), 70
        ws.add_image(img, "C1" if i == 0 else "D1")


def header_block(ws, first_col, groups, with_days):
    """groups: [(label, colours)] ; each spans 12 (with days) or 3 columns."""
    col = first_col
    span = 12 if with_days else 3
    for label, (dark, mid, light, txt) in groups:
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + span - 1)
        put(ws, 3, col, label, bg=dark, bold=True, color=txt, size=11)
        for i in range(span):
            ws.cell(3, col + i).fill = fill(dark)
            ws.cell(3, col + i).border = BOX
        subs = DAYS if with_days else [("WT", "Typical tournament day")]
        for d, (_, dname) in enumerate(subs):
            c0 = col + d * 3
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 2)
            put(ws, 4, c0, dname, bg=mid, bold=True, size=9, color="000000")
            for i in range(3):
                ws.cell(4, c0 + i).fill = fill(mid)
                ws.cell(4, c0 + i).border = BOX
                put(ws, 5, c0 + i, MEASURES[i][1], bg="FFFFFF", bold=True, size=8)
                put(ws, 7, c0 + i, MEASURES[i][2], bg=SECTION, bold=True, color="FFFFFF", size=7)
        col += span


def build_sheet(ws, groups, with_days, logos, title):
    data_first = 8
    n_data = 48 if with_days else 12
    tot_first = data_first + n_data
    last_col = tot_first + (12 if with_days else 0) - 1
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showZeros = False
    ws.sheet_view.zoomScale = 70
    banner(ws, last_col, logos, title)
    # descriptor headers
    for c, name, w in ((2, "Section", 4), (3, "Overlay", 16), (4, "Deployment Location", 16), (5, "Resource", 40),
                       (6, "Resource Origin", 19), (7, "Rank", 22)):
        ws.merge_cells(start_row=3, start_column=c, end_row=5, end_column=c)
        put(ws, 3, c, name, bg=NAVY, bold=True, color="FFFFFF", size=10)
        for r in (4, 5):
            ws.cell(r, c).fill = fill(NAVY)
        ws.column_dimensions[L(c)].width = w
    ws.column_dimensions["A"].width = 1.5
    ws.column_dimensions["B"].hidden = True
    header_block(ws, data_first, [(r, RISK_COL[r]) for r in RISKS], with_days)
    if with_days:
        ws.merge_cells(start_row=3, start_column=tot_first, end_row=3, end_column=last_col)
        put(ws, 3, tot_first, "Automatic Total Figures Calculation (all four days)", bg=TOTAL_COL[0], bold=True,
            color="FFFFFF", size=11)
        for i in range(12):
            ws.cell(3, tot_first + i).fill = fill(TOTAL_COL[0])
        for s, risk in enumerate(RISKS):
            c0 = tot_first + s * 3
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 2)
            put(ws, 4, c0, risk, bg=RISK_COL[risk][1], bold=True, size=9)
            for i in range(3):
                ws.cell(4, c0 + i).fill = fill(RISK_COL[risk][1])
                put(ws, 5, c0 + i, MEASURES[i][1], bg="FFFFFF", bold=True, size=8)
                put(ws, 7, c0 + i, MEASURES[i][2], bg=SECTION, bold=True, color="FFFFFF", size=7)
    for c in range(data_first, last_col + 1):
        ws.column_dimensions[L(c)].width = 9.5
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 44
    # row 6: hidden one-row header for Power BI
    names = ["Section", "Overlay", "Deployment Location", "Resource", "Resource Origin", "Rank"]
    for j, n in enumerate(names, 2):
        ws.cell(6, j, n)
    for s, risk in enumerate(RISKS):
        subs = DAYS if with_days else [("WT", "")]
        for d, (dcode, _) in enumerate(subs):
            for m, (mcode, _, _) in enumerate(MEASURES):
                ws.cell(6, data_first + s * len(subs) * 3 + d * 3 + m, f"{risk} | {dcode} | {mcode}")
        if with_days:
            for m, (mcode, _, _) in enumerate(MEASURES):
                ws.cell(6, tot_first + s * 3 + m, f"Total | {risk} | {mcode}")
    for c in range(2, last_col + 1):
        ws.cell(6, c).font = Font(name=FONT, size=7, color="808080")
    ws.row_dimensions[6].hidden = True
    # validations
    dv_off = DataValidation(type="whole", operator="between", formula1="0", formula2="9999", allow_blank=True,
                            showErrorMessage=True, errorTitle="Officers", error="Whole number of officers (0-9999).")
    dv_hrs = DataValidation(type="decimal", operator="between", formula1="0", formula2="24", allow_blank=True,
                            showErrorMessage=True, errorTitle="Hours", error="Hours per officer for the day (0-24).")
    dv_block = DataValidation(type="custom", formula1="FALSE", allow_blank=True, showErrorMessage=True,
                              errorTitle="Not required", error="Normal Tour of Duty is not required for mutual aid.")
    dv_origin = DataValidation(type="list", formula1=f'"{FS},{MA}"', allow_blank=True)
    dv_rank = DataValidation(type="list", formula1='"' + ",".join(RANKS) + '"', allow_blank=True)
    for dv in (dv_off, dv_hrs, dv_block, dv_origin, dv_rank):
        ws.add_data_validation(dv)
    # rows
    r = 7
    first_group = True
    for group in groups:
        if not first_group:
            r += 1                                            # spacer row
        section = group[0]["section"]
        # section bar (row 7 for the first group shares the units row)
        put(ws, r, 3, section, bg=SECTION, bold=True, color="FFFFFF", size=9)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        for c in range(3, 8):
            ws.cell(r, c).fill = fill(SECTION)
        if not first_group:
            for c in range(data_first, last_col + 1):
                ws.cell(r, c).fill = fill(SECTION)
        ws.cell(r, 2, section)
        first_group = False
        r += 1
        prev = None
        for row in group:
            ws.cell(r, 2, row["section"]).font = Font(name=FONT, size=7, color="808080")
            new_block = prev is None or (row["resource"], row["origin"], row["location"]) != \
                (prev["resource"], prev["origin"], prev["location"])
            vals = [(3, row["overlay"], "DDEBF7"), (4, row["location"], "DDEBF7"), (5, row["resource"], "F2F2F2"),
                    (6, row["origin"], "E2EFDA"), (7, row["rank"], "FCE4D6")]
            for c, v, bg in vals:
                cell = put(ws, r, c, v, bg=bg, size=8 if c != 5 else 9)
                if not new_block and c < 7 and not row["free"]:
                    cell.font = Font(name=FONT, size=8, color="A6A6A6")      # repeated label shown faint
            if row["free"]:
                for c in (5, 6, 7):
                    ws.cell(r, c).fill = fill("FFFFFF")
                dv_origin.add(f"F{r}")
                dv_rank.add(f"G{r}")
                ws.cell(r, 5).value = None
            for s, risk in enumerate(RISKS):
                tint = RISK_COL[risk][2]
                n_sub = 4 if with_days else 1
                for d in range(n_sub):
                    for m in range(3):
                        c = data_first + s * n_sub * 3 + d * 3 + m
                        cell = put(ws, r, c, bg=tint, size=9)
                        coord = f"{L(c)}{r}"
                        if m == 1 and row["origin"] == MA:
                            cell.fill = PatternFill(patternType="lightGray", fgColor="808080", bgColor=tint)
                            dv_block.add(coord)
                        elif m == 0:
                            cell.number_format = "0"
                            dv_off.add(coord)
                        else:
                            cell.number_format = "0.0"
                            dv_hrs.add(coord)
            if with_days:
                for s in range(4):
                    for m in range(3):
                        refs = ",".join(f"{L(data_first + s * 12 + d * 3 + m)}{r}" for d in range(4))
                        c = tot_first + s * 3 + m
                        put(ws, r, c, f"=SUM({refs})", bg=TOTAL_COL[2], size=9).number_format = "0.0" if m else "0"
            if new_block:
                for c in range(3, last_col + 1):
                    b = ws.cell(r, c).border
                    ws.cell(r, c).border = Border(left=b.left, right=b.right, bottom=b.bottom,
                                                  top=Side(style="medium", color="8EA9DB"))
            prev = row
            r += 1
    ws.freeze_panes = f"{L(data_first)}8"
    ws.oddHeader.center.text = ws.oddFooter.center.text = "OFFICIAL"
    ws.page_setup.orientation = "landscape"
    return r - 1


GUIDANCE = [
    "EURO 2028 - Resource Forecast Collection Form",
    "",
    "1. Fill in the Match Day Resource and Wider Tournament Resource sheets only.",
    "2. Each row is one resource, from one origin (Force Supplied or Mutual Aid Supplied), at one rank.",
    "3. Match Day Resource: for ONE fixture at each risk level (Low, Medium, High, Super High), enter for each day "
    "(Match Day -2, -1, Match Day, +1): Total Officers required, Normal Tour of Duty hours per officer, and Extra "
    "Hours / Over Time per officer.",
    "4. Wider Tournament Resource: the same, for a typical tournament day with no fixture in your area.",
    "5. Hatched cells are not required (Normal Tour of Duty for mutual aid). Leave anything you do not need blank.",
    "6. Additional Resources: type the resource, then choose the origin and rank from the lists.",
    "7. Do not insert, delete or move rows or columns, and do not rename the sheets.",
    "8. Save the completed form as your force name, e.g. 'West Midlands Police.xlsx'.",
]


def build(out):
    z = zipfile.ZipFile(BLANK)
    logos = [z.read("xl/media/image4.jpg"), z.read("xl/media/image5.png")]
    wb = Workbook()
    g = wb.active
    g.title = "Guidance"
    g.sheet_view.showGridLines = False
    g.column_dimensions["B"].width = 140
    for i, t in enumerate(GUIDANCE, 2):
        c = g.cell(i, 2, t)
        c.font = Font(name=FONT, size=16 if i == 2 else 11, bold=i == 2, color=NAVY if i == 2 else "000000")
        c.alignment = Alignment(wrap_text=True, vertical="top")
    md = wb.create_sheet("Match Day Resource")
    md_last = build_sheet(md, md_rows(), True, logos, "Euro2028 Resource Forecast Collection Form - Match Day")
    wt = wb.create_sheet("Wider Tournament Resource")
    wt_last = build_sheet(wt, wt_rows(), False, logos, "Euro2028 Resource Forecast Collection Form - Wider Tournament")
    wb.active = 1
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out, md_last, wt_last


if __name__ == "__main__":
    print(build(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
