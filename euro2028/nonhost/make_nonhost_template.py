"""
EURO 2028 Non-Host Resource Collection Form - sent to non-host forces when an operation
(e.g. Op Brock, Op Stack) is stood up, so the resources it needs can be set against the
resources committed to EURO 2028.

Sheets
  Guidance             purpose, operation details (input), instructions, definitions, checklist
  Op Data Collection   one row per resource / rank: Number of officers, Opportunity Cost and
                       Chargeable Hours (hours per officer), with completeness checks
  Lists (hidden)       Skill_Portfolio / Skill list for the Resource drop-down (paste the full
                       MA_Skills_Portfolio list over it) and the rank list

Power BI reads: tblOpDetails (Guidance) and tblOpResources (Op Data Collection).

Usage: python3 make_nonhost_template.py [out.xlsx]
"""
import io
import os
import sys
import zipfile

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
BLANK = os.path.join(HERE, "..", "v1_lift_shift", "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "output", "EURO_2028_NonHost_Resource_Collection_Form.xlsx")
FONT = "Aptos Narrow"
NAVY, BLUE, MIDBLUE, LIGHT = "1F3864", "2F5597", "8EA9DB", "DDEBF7"
YELLOW, RED, GREY, WHITE, INPUT = "FFFF00", "C00000", "F2F2F2", "FFFFFF", "FFFFFF"
thin = Side(style="thin", color="A6A6A6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
ROWS = 150
RANKS = ["Assistant Chief Constable", "Chief Superintendent", "Superintendent", "Chief Inspector", "Inspector",
         "Sergeant", "Constable"]
# Skills read from the MA_Skills_Portfolio list (Armed Policing shown); paste the full list into 'Lists'
SKILLS = [("Armed Policing", s) for s in [
    ".338 Rifle Officer", "Armed Response to Public Disorder", "Armed Response Vehicle (ARV) Officer",
    "Armed Rural Surveillance Officer", "Armed Support to Public Disorder", "Armed Surveillance Officer",
    "ARV CBRN Officer", "Authorised Firearms Officer (AFO)", "Authorised Firearms Officer (CT)",
    "Basic Mobile Armed Support To Surveillance (MASTS)", "Cadre Tactical Firearms Commander - Silver",
    "Counter Assault Team (CAT)", "Counter Assault Team (CAT) Critical National Infrastructure",
    "Counter Terrorism Specialist Firearms Officer (CT SFO)", "Dynamic Search Officer", "Firearms Tactical Advisor",
    "High Profile Vehicle Escort Driver", "High Profile Vehicle Escort Officer", "Initial Tactical Firearms Commander",
    "National Firearms Instructor", "Operational AEP Commander", "Operational Firearms Commander", "Police Armourer",
    "Police Weapon Issuing Officer", "Police Weapons Maintainer", "Post Incident Manager (PIM)", "Protection Officer",
    "Protection Officer Driver", "Public Order AEP Operator (AFO)", "Public Order AEP Operator (Non-AFO)",
    "Rifle Officer", "Specialist Firearms Officer", "Specialist Strategic Firearms Commander",
    "Specialist Tactical Firearms Commander"]]


def fill(c):
    return PatternFill("solid", fgColor=c)


def font(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def banner(ws, last_col, title, logos, title_col=3):
    for r in range(1, 8):
        for c in range(1, last_col + 1):
            ws.cell(r, c).fill = fill(NAVY if r < 7 else MIDBLUE)
    for r, h in zip(range(1, 8), [8, 30, 4, 18, 8, 8, 5]):
        ws.row_dimensions[r].height = h
    t = ws.cell(2, title_col, title)
    t.font = font(20, True, "FFFFFF")
    t.alignment = Alignment(vertical="center")
    s = ws.cell(4, title_col, "National Police Coordination Centre - EURO 2028 Police Planning")
    s.font = font(11, False, "DDEBF7")
    off = ws.cell(2, last_col, "OFFICIAL")
    off.font = font(12, True, "FFFFFF")
    off.alignment = Alignment(horizontal="right")
    for i, data in enumerate(logos):
        img = XLImage(io.BytesIO(data))
        ratio = 52 / img.height
        img.width, img.height = int(img.width * ratio), 52
        ws.add_image(img, "A2" if i == 0 else ("B2" if title_col == 3 else "A2"))
        if title_col != 3 and i == 1:
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
            from openpyxl.drawing.xdr import XDRPositiveSize2D
            from openpyxl.utils.units import pixels_to_EMU
            img.anchor = OneCellAnchor(_from=AnchorMarker(col=0, colOff=pixels_to_EMU(78), row=1, rowOff=0),
                                       ext=XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height)))


def guidance(wb, logos):
    ws = wb.active
    ws.title = "Guidance"
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [14, 14, 36, 62, 30]):
        ws.column_dimensions[col].width = w
    banner(ws, 5, "Non-Host Resource Collection Form", logos)
    r = 9

    def heading(text):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, text)
        c.font = font(12, True)
        for cc in range(1, 6):
            ws.cell(r, cc).fill = fill(YELLOW)
        ws.row_dimensions[r].height = 20
        r += 1

    def text(t, color="000000", bold=False, chars=150):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        c = ws.cell(r, 1, t)
        c.font = font(11, bold, color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 16 * (len(t) // chars + 1)
        r += 1

    heading("Overview:")
    text("Where a national or regional operation (for example Op Brock or Op Stack) is stood up during the EURO 2028 "
         "tournament period, the officers it needs are not available for EURO 2028 - including as mutual aid. This "
         "form collects the resources your force is committing to that operation, so the NPoCC EURO 2028 Police "
         "Planning Team can see the combined demand on policing and identify any pressure on tournament resourcing.")
    text("Returns are combined in the national EURO 2028 Power BI dashboard alongside EURO 2028 demand, by date and "
         "skill.")
    r += 1
    heading("Operation details (please complete):")
    ws.cell(r, 3, "Field").font = font(11, True, "FFFFFF")
    ws.cell(r, 4, "Your entry").font = font(11, True, "FFFFFF")
    for cc in (3, 4):
        ws.cell(r, cc).fill = fill(NAVY)
        ws.cell(r, cc).border = BOX
    top = r
    r += 1
    fields = [("Operation name", "e.g. Op Brock"), ("Force", "Your force name"),
              ("Operation start date", "dd/mm/yyyy"), ("Operation end date", "dd/mm/yyyy"),
              ("Completed by", "Name and role"), ("Contact email", ""), ("Contact phone", ""),
              ("Date completed", "dd/mm/yyyy")]
    first_input = r
    for name, hint in fields:
        a = ws.cell(r, 3, name)
        a.font = font(11, True)
        a.fill = fill(LIGHT)
        a.border = BOX
        b = ws.cell(r, 4)
        b.fill = fill("FFF2CC")
        b.border = BOX
        b.font = font(11)
        b.protection = Protection(locked=False)
        if "date" in name.lower():
            b.number_format = "dd/mm/yyyy"
        if hint:
            h = ws.cell(r, 5, hint)
            h.font = font(9, italic=True, color="808080")
        ws.row_dimensions[r].height = 20
        r += 1
    t = Table(displayName="tblOpDetails", ref=f"C{top}:D{r - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)
    dv = DataValidation(type="date", operator="between", formula1="46023", formula2="47848", allow_blank=True,
                        showErrorMessage=True, errorTitle="Date", error="Enter a date between 2026 and 2030, "
                                                                        "e.g. 14/06/2028.")
    dv.add(f"D{first_input + 2}:D{first_input + 3}")
    dv.add(f"D{first_input + 7}")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="custom", formula1=f"D{first_input + 3}>=D{first_input + 2}", allow_blank=True,
                         showErrorMessage=True, errorTitle="End date", error="The end date must be on or after the "
                                                                              "start date.")
    dv2.add(f"D{first_input + 3}")
    ws.add_data_validation(dv2)
    r += 1
    heading("Instructions for forces:")
    for i, t_ in enumerate([
            "Complete the Operation details above - the operation name and the dates it is stood up are essential.",
            "Go to the 'Op Data Collection' tab. Use one row per resource and rank.",
            "Choose the Resource from the drop-down list; the Skill Portfolio fills in automatically.",
            "Choose the Rank, then enter the Number of Officers, the Opportunity Cost and the Chargeable Hours.",
            "Opportunity Cost and Chargeable Hours are hours PER OFFICER PER DAY. Enter 0 if there are none - do not "
            "leave them blank, or the line cannot be costed.",
            "Check the status box at the top of the data tab reads 'READY TO SUBMIT' before you send the form.",
            "Do not add, delete, move or rename rows, columns or tabs. Save the file as "
            "'<Force name> - <Operation name>.xlsx'."], 1):
        text(f"{i}. {t_}")
    text("Returns should be provided by [DATE / TIME - TBC] to [EMAIL ADDRESS - TBC].", RED, True)
    r += 1
    heading("Important notes and guidance:")
    for title, t_ in [
            ("Resources", "The Resource list is the national mutual aid skills list, so your return matches the "
                          "skills used in EURO 2028 planning. If a skill is not listed, choose the closest and explain "
                          "it in Notes."),
            ("Number of Officers", "Officers needed per day for the operation. If the number changes during the "
                                   "operation, give the peak daily figure and explain in Notes."),
            ("Opportunity Cost", "Normal rostered (tour of duty) hours per officer per day - the time officers are "
                                 "taken away from other duties."),
            ("Chargeable Hours", "Extra / overtime hours per officer per day. Enter 0 if none."),
            ("Updates", "If the operation or your commitment changes, send an updated form with the same operation "
                        "name - the latest version will be used."),
            ("Handling", "This form and your return are OFFICIAL.")]:
        a = ws.cell(r, 1, title)
        a.font = font(11, True)
        a.alignment = Alignment(vertical="top", wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        b = ws.cell(r, 3, t_)
        b.font = font(11)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 16 * (len(t_) // 115 + 1)
        r += 1
    r += 1
    heading("Before you submit - checklist:")
    for t_ in ["Operation name, start date and end date completed.",
               "Every row used has a Resource, Rank, Number of Officers, Opportunity Cost and Chargeable Hours.",
               "Status box on the data tab reads READY TO SUBMIT.",
               "File saved as '<Force name> - <Operation name>.xlsx' and sent by the deadline."]:
        text("[  ]  " + t_)
    r += 1
    heading("Contacts:")
    text("NPoCC EURO 2028 Police Planning Team - [NAME / EMAIL / PHONE - TBC]", RED)
    ws.oddHeader.center.text = ws.oddFooter.center.text = "OFFICIAL"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    return ws, first_input


def data_sheet(wb, logos, op_row):
    ws = wb.create_sheet("Op Data Collection")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    cols = [("Skill Portfolio", 22), ("Resource", 52), ("Rank", 24), ("Number of Officers", 14),
            ("Opportunity Cost\n(hrs per officer)", 16), ("Chargeable Hours\n(hrs per officer)", 16),
            ("Notes", 40), ("Status", 18)]
    for j, (_, w) in enumerate(cols, 1):
        ws.column_dimensions[L(j)].width = w
    banner(ws, len(cols), "Non-Host Resource Collection Form", logos, title_col=2)
    # operation summary strip (from Guidance)
    labels = [("Operation:", f"=IF(Guidance!D{op_row}=\"\",\"Complete on Guidance tab\",Guidance!D{op_row})"),
              ("Force:", f"=IF(Guidance!D{op_row + 1}=\"\",\"Complete on Guidance tab\",Guidance!D{op_row + 1})"),
              ("From:", f"=IF(Guidance!D{op_row + 2}=\"\",\"-\",TEXT(Guidance!D{op_row + 2},\"dd/mm/yyyy\"))"),
              ("To:", f"=IF(Guidance!D{op_row + 3}=\"\",\"-\",TEXT(Guidance!D{op_row + 3},\"dd/mm/yyyy\"))")]
    r = 9
    for i, (lab, f) in enumerate(labels):
        c = (i % 2) * 2 + 1
        rr = r + i // 2
        a = ws.cell(rr, c, lab)
        a.font = font(11, True, NAVY)
        a.alignment = Alignment(horizontal="right")
        b = ws.cell(rr, c + 1, f)
        b.font = font(11, True)
        b.fill = fill(LIGHT)
        b.border = BOX
        if c + 1 == 4:
            ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=6)
            for cc in (5, 6):
                ws.cell(rr, cc).fill = fill(LIGHT)
                ws.cell(rr, cc).border = BOX
    hdr = 14
    last = hdr + ROWS
    # completion panel
    ws.merge_cells("G9:H9")
    ws["G9"] = "Completion check"
    ws["G9"].font = font(11, True, "FFFFFF")
    ws["G9"].fill = fill(NAVY)
    ws["G9"].alignment = Alignment(horizontal="center")
    ws.merge_cells("G10:H10")
    ws["G10"] = (f'=IF(COUNTA(B{hdr + 1}:B{last})=0,"No resources entered yet",'
                 f'IF(COUNTIF(H{hdr + 1}:H{last},"Check*")+(Guidance!D{op_row}="")+(Guidance!D{op_row + 2}="")'
                 f'+(Guidance!D{op_row + 3}="")=0,"READY TO SUBMIT",'
                 f'COUNTIF(H{hdr + 1}:H{last},"Check*")+(Guidance!D{op_row}="")+(Guidance!D{op_row + 2}="")'
                 f'+(Guidance!D{op_row + 3}="")&" item(s) to complete"))')
    ws["G10"].font = font(12, True)
    ws["G10"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[10].height = 24
    for c in ("G10", "H10"):
        ws[c].border = BOX
    ws.conditional_formatting.add("G10:H10", FormulaRule(formula=['$G$10="READY TO SUBMIT"'],
                                                         fill=fill("C6EFCE"), font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add("G10:H10", FormulaRule(formula=['ISNUMBER(SEARCH("to complete",$G$10))'],
                                                         fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.merge_cells(start_row=12, start_column=1, end_row=12, end_column=len(cols))
    ws.cell(12, 1, "Resources required for the operation (one row per resource and rank). Yellow = choose from "
                   "list, white = type, grey = filled automatically.").font = font(10, italic=True, color="404040")
    # header
    for j, (name, _) in enumerate(cols, 1):
        c = ws.cell(hdr, j, name)
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[hdr].height = 46
    for r in range(hdr + 1, last + 1):
        band = GREY if (r - hdr) % 2 == 0 else WHITE
        formulas = {1: f'=IF(B{r}="","",IFERROR(INDEX(Lists!$A:$A,MATCH(B{r},Lists!$B:$B,0)),"Not on list"))',
                    8: (f'=IF(AND(B{r}="",C{r}="",D{r}="",E{r}="",F{r}=""),"",'
                        f'IF(OR(B{r}="",C{r}="",D{r}="",E{r}="",F{r}=""),"Check - missing","OK"))')}
        for j in range(1, len(cols) + 1):
            c = ws.cell(r, j, formulas.get(j))
            c.border = BOX
            c.font = font(10)
            c.alignment = Alignment(vertical="center", horizontal="center" if j in (4, 5, 6, 8) else "left",
                                    wrap_text=j in (2, 7))
            if j in (1, 8):
                c.fill = fill("E7E6E6")
            elif j in (2, 3):
                c.fill = fill("FFF2CC")
                c.protection = Protection(locked=False)
            else:
                c.fill = fill(band)
                c.protection = Protection(locked=False)
            if j == 4:
                c.number_format = "0"
            if j in (5, 6):
                c.number_format = "0.0"
    t = Table(displayName="tblOpResources", ref=f"A{hdr}:{L(len(cols))}{last}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)
    rng = lambda col: f"{col}{hdr + 1}:{col}{last}"  # noqa: E731
    dvs = [DataValidation(type="list", formula1="=SkillList", allow_blank=True, showErrorMessage=True,
                          errorTitle="Resource", error="Choose a resource from the list.", promptTitle="Resource",
                          prompt="Choose from the list.", showInputMessage=True),
           DataValidation(type="list", formula1="=RankList", allow_blank=True, showErrorMessage=True,
                          errorTitle="Rank", error="Choose a rank from the list."),
           DataValidation(type="whole", operator="between", formula1="0", formula2="9999", allow_blank=True,
                          showErrorMessage=True, errorTitle="Officers", error="Whole number of officers (0-9999)."),
           DataValidation(type="decimal", operator="between", formula1="0", formula2="24", allow_blank=True,
                          showErrorMessage=True, errorTitle="Hours", error="Hours per officer per day (0-24).")]
    for dv, cl in zip(dvs, ["B", "C", "D", "E"]):
        ws.add_data_validation(dv)
        dv.add(rng(cl))
    dvs[3].add(rng("F"))
    # completeness highlighting
    body = f"B{hdr + 1}:F{last}"
    ws.conditional_formatting.add(body, FormulaRule(
        formula=[f'AND(B{hdr + 1}="",COUNTA($B{hdr + 1}:$F{hdr + 1})>0)'], fill=fill("FFC7CE")))
    ws.conditional_formatting.add(rng("H"), FormulaRule(formula=[f'LEFT(H{hdr + 1},5)="Check"'],
                                                        fill=fill("FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(rng("H"), FormulaRule(formula=[f'H{hdr + 1}="OK"'], fill=fill("C6EFCE"),
                                                        font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(rng("A"), FormulaRule(formula=[f'A{hdr + 1}="Not on list"'],
                                                        font=Font(color="9C0006", bold=True)))
    ws.freeze_panes = f"A{hdr + 1}"
    ws.oddHeader.center.text = ws.oddFooter.center.text = "OFFICIAL"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hdr}:{hdr}"
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False
    return ws


def lists(wb):
    ws = wb.create_sheet("Lists")
    ws["A1"], ws["B1"], ws["D1"] = "Skill_Portfolio", "Skill", "Rank"
    for c in ("A1", "B1", "D1"):
        ws[c].font = font(10, True)
    ws["F1"] = ("Paste the full MA_Skills_Portfolio list into columns A:B from row 2 (Skill_Portfolio, Skill). The "
                "Resource drop-down reads B2:B2000.")
    ws["F1"].font = font(10, True, RED)
    for i, (p, s) in enumerate(SKILLS, 2):
        ws.cell(i, 1, p)
        ws.cell(i, 2, s)
    for i, rk in enumerate(RANKS, 2):
        ws.cell(i, 4, rk)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 24, 55
    ws.column_dimensions["D"].width = 26
    wb.defined_names["SkillList"] = DefinedName("SkillList", attr_text="Lists!$B$2:$B$2000")
    wb.defined_names["RankList"] = DefinedName("RankList", attr_text=f"Lists!$D$2:$D${len(RANKS) + 1}")
    ws.sheet_state = "hidden"


def build(out=OUT):
    z = zipfile.ZipFile(BLANK)
    logos = [z.read("xl/media/image4.jpg"), z.read("xl/media/image5.png")]
    wb = Workbook()
    g, op_row = guidance(wb, logos)
    data_sheet(wb, logos, op_row)
    lists(wb)
    wb.active = 0
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(build(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
