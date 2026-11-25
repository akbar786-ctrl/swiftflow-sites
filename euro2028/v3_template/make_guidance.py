"""
Guidance sheet for the EURO 2028 Resource Collection Template V2, in the style of the
national team's previous collections (yellow section bars, definitions table).
Produces a one-sheet workbook; copy the sheet into the template with
Right-click tab > Move or Copy > (choose the template) > tick Create a copy.

Usage: python3 make_guidance.py [out.xlsx]
"""
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output", "EURO_2028_Guidance_Sheet.xlsx")
NAVY, YELLOW, RED, GREY = "1F3864", "FFFF00", "C00000", "F2F2F2"
FONT = "Aptos Narrow"
thin = Side(style="thin", color="A6A6A6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

OVERVIEW = [
    "UEFA EURO 2028 will be hosted across the United Kingdom and Ireland. The National Police Coordination Centre "
    "(NPoCC) EURO 2028 Police Planning Team is building a national picture of the policing resources the tournament "
    "will require, and their cost.",
    "This form collects each force's forecast of the resources it needs - by resource, rank and origin (force supplied "
    "or mutual aid) - for match days and for the wider tournament. Returns from all forces are combined in a national "
    "Power BI dashboard used for capacity, mutual aid and budget planning at force, regional and national level.",
    "Fixture allocations and competing nations are not yet confirmed, so you are asked to complete every risk level "
    "(Low, Medium, High and Super High). The planning team will apply the appropriate risk level to each fixture.",
]
INSTRUCTIONS = [
    "Complete the Match Day Resource and Wider Tournament Resource tabs. The Metropolitan Police Service should also "
    "complete the Pan London tab.",
    "Match Day Resource: what your force needs for ONE fixture at each risk level, for each of Match Day -2, Match "
    "Day -1, Match Day and Match Day +1.",
    "Wider Tournament Resource: what your force needs on a TYPICAL tournament day when there is no fixture in your "
    "area, at each risk level.",
    "For every line you use, complete all three columns: Total Officers required, Normal Tour of Duty (hours per "
    "officer) and Chargeable Hours (hours per officer). Costs are calculated from the hours - if hours are blank, no "
    "cost is counted for that line.",
    "Only type in the white cells. Grey hatched cells are not required for that resource or day and are locked.",
    "Do not insert, delete, move or rename rows, columns or tabs - the returns are read automatically and any change "
    "to the layout will stop your figures being counted.",
    "Save the completed form as your force name, e.g. 'West Midlands Police.xlsx'.",
]
DEADLINE = "Returns should be provided by [DATE / TIME - TBC] to [EMAIL ADDRESS - TBC]."
NOTES = [
    ("PSU / Serial Configurations (top of each tab)",
     "Enter how many officers of each rank make up ONE unit - for example a Level 2 PSU of 1 Inspector, 3 Sergeants "
     "and 21 Constables is entered as 1, 3 and 21. Then, in the Public Order / Public Safety section, enter the NUMBER "
     "of PSUs or BDUs required. Officers are calculated as number of units x officers per unit, so both parts must "
     "be completed. Configurations left blank mean your PSUs cannot be counted or costed."),
    ("Match Day only and Non-Match Day resources",
     "Resources marked (Match Day) or (Match Day Only) are completed in the Match Day columns only. Resources marked "
     "(Non-Match Day) are completed in the Match Day -2, Match Day -1 and Match Day +1 columns only. The columns that "
     "do not apply are hatched."),
    ("Blank or zero",
     "Leave a line blank if you do not need that resource. If you need officers but no chargeable (extra) hours, "
     "enter 0 - do not leave it blank - so we know it has not been missed."),
    ("Hours", "Normal Tour of Duty and Chargeable Hours are hours PER OFFICER for that day, not totals for the team "
              "(e.g. 20 officers each working 10 hours with 2 extra hours: 20, 10, 2)."),
    ("Mutual aid", "Mutual Aid Requested lines are the resources you need from other forces. Normal Tour of Duty is "
                   "not required for mutual aid and is hatched."),
    ("Ranks and roles", "Each line is one rank. Where a resource is a role or skill (e.g. Bronze Commander, CBRN "
                        "Search), use the line provided for it - the rank is already set."),
    ("Additional Resources", "Use these lines for anything not listed: describe the resource and choose the origin "
                             "and rank from the drop-down lists."),
    ("Risk levels", "Low, Medium, High and Super High are the assumed risk of a fixture or tournament day. Complete "
                    "every level, even if your current planning assumption is a single level."),
    ("Notes", "If you need to explain an assumption, add it in the Notes column (or contact the team) - do not type "
              "notes in the number cells."),
    ("Handling", "This form and your return are OFFICIAL. Data will be shared with the NPoCC EURO 2028 Police "
                 "Planning Team and partners involved in tournament planning."),
]
FIELDS = [
    ("Overlay", "Whether the resource is an overlay (e.g. Armed Response Overlay).", "Pre-filled"),
    ("Deployment Location", "Where the resource is deployed: In Stadium, Stadium Vicinity, Town Centre / Any other "
                            "location; Wider Community or Fan Zone on the Wider Tournament tab.", "Pre-filled"),
    ("Resource", "The resource or role required.", "Pre-filled (type it only on Additional Resources lines)"),
    ("Resource Origin", "Force Supplied = your own officers. Mutual Aid Requested = officers you need from other "
                        "forces.", "Pre-filled (choose on Additional Resources lines)"),
    ("Rank", "The rank of the officers on that line (Inspector, Sergeant, Constable ...). On PSU lines: the unit "
             "type (Level 2 PSU / Level 3 BDU).", "Pre-filled (choose on Additional Resources lines)"),
    ("Total Officers required", "Number of officers (or, on PSU / BDU lines, number of units) needed that day.",
     "ESSENTIAL - whole number"),
    ("Normal Tour of Duty", "Normal rostered hours per officer for that day.", "ESSENTIAL - hours per officer "
                                                                           "(force supplied)"),
    ("Chargeable Hours", "Extra / overtime hours per officer for that day. Drives the cost.", "ESSENTIAL - hours "
                                                                                             "per officer; 0 if none"),
    ("Match Day -2 / -1 / Match Day / +1", "Two days before, the day before, the day of and the day after the "
                                           "fixture.", "Match Day tab"),
    ("Low / Medium / High / Super High", "Risk level of the fixture or tournament day.", "Complete every level"),
    ("PSU / Serial Configurations", "Officers of each rank in ONE PSU / BDU (e.g. 1 / 3 / 21).",
     "ESSENTIAL if you request any PSUs / BDUs"),
]
CHECKLIST = [
    "Every risk level (Low, Medium, High, Super High) completed.",
    "PSU / Serial Configurations completed for every PSU / BDU type you have requested.",
    "Every line with officers also has Normal Tour of Duty and Chargeable Hours (0 if no extra hours).",
    "Nothing typed in hatched cells; no rows, columns or tabs added, removed or renamed.",
    "Additional Resources lines have a description, origin and rank.",
    "File saved as your force name and sent by the deadline.",
]


def build(out=OUT):
    wb = Workbook()
    ws = wb.active
    ws.title = "Guidance"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 42
    r = 1

    def fill_row(row, colour):
        for c in range(2, 5):
            ws.cell(row, c).fill = PatternFill("solid", fgColor=colour)

    ws.merge_cells("B1:D1")
    ws["B1"] = "EURO 2028 Resource Forecast Collection - Instructions and Guidance"
    ws["B1"].font = Font(name=FONT, size=16, bold=True, color="FFFFFF")
    ws["B1"].alignment = Alignment(vertical="center")
    fill_row(1, NAVY)
    ws.row_dimensions[1].height = 34
    ws["D2"] = "OFFICIAL"
    ws["D2"].font = Font(name=FONT, size=11, bold=True, color=RED)
    ws["D2"].alignment = Alignment(horizontal="right")
    r = 3

    def heading(text):
        nonlocal r
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(r, 2, text)
        c.font = Font(name=FONT, size=12, bold=True)
        fill_row(r, YELLOW)
        r += 1

    def text(t, bold=False, color="000000", height=None):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(r, 2, t)
        c.font = Font(name=FONT, size=11, bold=bold, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = height or max(16, 15 * (len(t) // 150 + 1))
        r += 1

    heading("Overview:")
    for t in OVERVIEW:
        text(t)
    heading("Instructions for forces:")
    for i, t in enumerate(INSTRUCTIONS, 1):
        text(f"{i}. {t}")
    text(DEADLINE, bold=True, color=RED)
    heading("Important notes and guidance:")
    for title, t in NOTES:
        c = ws.cell(r, 2, title)
        c.font = Font(name=FONT, size=11, bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        c = ws.cell(r, 3, t)
        c.font = Font(name=FONT, size=11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = max(16, 15 * (len(t) // 105 + 1))
        r += 1
    heading("Column definitions:")
    for j, h in enumerate(["Field", "Definition", "What to enter"], 2):
        c = ws.cell(r, j, h)
        c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.border = BOX
    r += 1
    for f, d, w in FIELDS:
        for j, v in enumerate([f, d, w], 2):
            c = ws.cell(r, j, v)
            c.font = Font(name=FONT, size=10.5, bold=(j == 2 or "ESSENTIAL" in v and j == 4),
                          color=RED if (j == 4 and "ESSENTIAL" in v) else "000000")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOX
            c.fill = PatternFill("solid", fgColor=GREY)
        ws.row_dimensions[r].height = max(16, 15 * max(len(d) // 70 + 1, len(w) // 38 + 1, len(f) // 34 + 1))
        r += 1
    heading("Colour key:")
    for label, colour, pattern in (("Type here", "FFFFFF", None), ("Not required - locked", "D9D9D9", "lightUp"),
                                   ("Section heading - no entry", NAVY, None)):
        c = ws.cell(r, 2)
        c.fill = PatternFill(patternType=pattern or "solid", fgColor="808080" if pattern else colour,
                             bgColor=colour if pattern else None) if pattern else PatternFill("solid", fgColor=colour)
        c.border = BOX
        ws.cell(r, 3, label).font = Font(name=FONT, size=11)
        r += 1
    heading("Before you submit - checklist:")
    for t in CHECKLIST:
        text("[  ]  " + t)
    heading("Contacts:")
    text("NPoCC EURO 2028 Police Planning Team - [NAME / EMAIL / PHONE - TBC]", color=RED)
    ws.oddHeader.center.text = ws.oddFooter.center.text = "OFFICIAL"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(build(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
