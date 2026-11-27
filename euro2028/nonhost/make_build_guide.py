"""Step-by-step guide for finishing the team's own Non-Host form by hand in Excel (Word)."""
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "v3_template"))
from make_rules_doc import bullets, formula, para, table  # noqa: E402

OUT = os.path.join(HERE, "output", "EURO_2028_NonHost_Form_Build_Guide.docx")
NAVY = RGBColor(0x1F, 0x38, 0x64)


def steps(doc, items):
    n = 0
    for it in items:
        if isinstance(it, tuple) and it[0] == "f":
            formula(doc, it[1])
            continue
        n += 1
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.6)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        p.add_run(f"{n}.  ").bold = True
        p.add_run(it)


def part(doc, title, intro=None):
    doc.add_heading(title, level=1)
    if intro:
        para(doc, intro, italic=True)


def build(out=OUT):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10)
    for n in ("Heading 1", "Heading 2"):
        doc.styles[n].font.name = "Arial"
        doc.styles[n].font.color.rgb = NAVY
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.8)
    for p_ in (sec.header.paragraphs[0], sec.footer.paragraphs[0]):
        p_.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_.add_run("OFFICIAL")
        r.bold, r.font.color.rgb = True, RGBColor(0x1F, 0x4E, 0x9A)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("EURO 2028 Non-Host Resource Collection Form\nStep-by-step build guide")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(18), NAVY
    para(doc, "For finishing your own Non-Host form in Excel. It assumes your current layout: Operation box D11, "
              "Force box beside 'Force:', From box D13, To box beside 'To:', and Table1 with headers in row 16 "
              "(C Skill Portfolio, D Resource, E Rank, F Number of Officers, G Opportunity Cost, H Chargeable, "
              "I Notes). Work through the parts in order and save after each one.", italic=True)

    table(doc, ["Part", "What it does", "Time"], [
        ["1", "Lists tab for the drop-downs", "10 min"],
        ["2", "Names for the operation detail boxes", "5 min"],
        ["3", "Table size, headers and Status column", "5 min"],
        ["4", "Automatic formulas", "5 min"],
        ["5", "Drop-downs and entry rules", "15 min"],
        ["6", "Red missing-cell highlighting", "10 min"],
        ["7", "Completion check box", "5 min"],
        ["8", "Guidance tab text", "30-40 min"],
        ["9", "Lock, protect and hide", "10 min"],
        ["10", "Test", "10 min"],
    ], [1.2, 12, 3.8])

    # 1
    part(doc, "Part 1 - Lists tab", "Holds the Resource and Rank lists the drop-downs read from.")
    steps(doc, [
        "Click the + next to the sheet tabs to add a sheet. Right-click it › Rename › type Lists.",
        "Type the headings: A1 Skill Portfolio, B1 Skill, D1 Rank.",
        "Paste the MA skills list from row 2: portfolio in column A, skill in column B (one skill per row, no blank "
        "rows). Use Paste Special › Values if it comes from another file.",
        "Type the ranks in D2 downwards: Assistant Chief Constable, Chief Superintendent, Superintendent, Chief "
        "Inspector, Inspector, Sergeant, Constable (add Detective ranks or Police Staff if you want them).",
        "Formulas › Name Manager › New. Name: SkillList. Refers to:",
        ("f", "=OFFSET(Lists!$B$2,0,0,COUNTA(Lists!$B:$B)-1,1)"),
        "New again. Name: RankList. Refers to:",
        ("f", "=OFFSET(Lists!$D$2,0,0,COUNTA(Lists!$D:$D)-1,1)"),
        "Close Name Manager. (These grow automatically if you add to the lists, and drop-downs show no blank lines.)",
    ])

    # 2
    part(doc, "Part 2 - Name the operation detail boxes",
         "Naming the boxes means the formulas later work whatever cell they are in.")
    steps(doc, [
        "Go to the Op Data Collection tab. Click the box where the operation name is typed (D11).",
        "Click in the Name Box (left of the formula bar), type OpName and press Enter.",
        "Click the box beside 'Force:' › Name Box › OpForce › Enter.",
        "Click the box beside 'From:' (D13) › Name Box › OpStart › Enter.",
        "Click the box beside 'To:' › Name Box › OpEnd › Enter.",
        "Check in Formulas › Name Manager that all four point at the right cells.",
    ])

    # 3
    part(doc, "Part 3 - Table size, headers and Status column")
    steps(doc, [
        "Click J16 and type Status. The table extends to include it.",
        "Click H16 and change the header to: Chargeable Hours (Hours per Officer) - so it matches the Host template "
        "and the Guidance.",
        "Click inside the table › Table Design › Resize Table › type =$C$16:$J$116 › OK. This gives 100 rows. "
        "(A protected sheet cannot grow the table, so make it big enough now.)",
    ])

    # 4
    part(doc, "Part 4 - Automatic formulas",
         "Type each in row 17 only - the table copies it down every row by itself.")
    steps(doc, [
        "C17 (Skill Portfolio - fills in from the Resource chosen):",
        ("f", '=IF(D17="","",IFERROR(INDEX(Lists!$A:$A,MATCH(D17,Lists!$B:$B,0)),"Not on list"))'),
        "J17 (Status - blank for an unused row, 'Check - missing' if any of D to H is empty, otherwise 'OK'):",
        ("f", '=IF(COUNTA(D17:H17)=0,"",IF(COUNTBLANK(D17:H17)>0,"Check - missing","OK"))'),
        "If a small 'AutoCorrect Options' icon appears, leave it - you want the column filled.",
        "Shade C and J light grey (Home › Fill Color) so forces can see they are automatic.",
    ])

    # 5
    part(doc, "Part 5 - Drop-downs and entry rules", "All under Data › Data Validation. For each: select the "
                                                         "cells, set the Settings tab, then the Error Alert tab.")
    table(doc, ["Cells", "Allow / settings", "Error message"], [
        ["D17:D116 Resource", "List. Source: =SkillList", "Choose a resource from the list."],
        ["E17:E116 Rank", "List. Source: =RankList", "Choose a rank from the list."],
        ["F17:F116 Officers", "Whole number, between 0 and 9999", "Enter a whole number of officers."],
        ["G17:H116 Hours", "Decimal, between 0 and 24", "Hours per officer per day (0 to 24)."],
        ["OpStart (From)", "Date, between 01/01/2028 and 31/12/2028", "Enter a date, e.g. 14/06/2028."],
        ["OpEnd (To)", "Custom. Formula: =AND(ISNUMBER(OpEnd),OpEnd>=OpStart)",
         "Enter a date on or after the start date."],
    ], [3.6, 7.2, 6.2])
    bullets(doc, [
        "Leave 'Ignore blank' ticked on all of them and Style: Stop.",
        "Format OpStart and OpEnd as dates: Home › Number Format › Short Date (dd/mm/yyyy).",
        "Optional: on the Input Message tab of D17:D116 add 'Choose the resource from the list - the Skill Portfolio "
        "fills in automatically.'",
    ])

    # 6
    part(doc, "Part 6 - Red missing-cell highlighting",
         "Same idea as the Host template: once any cell on a row is filled, the empty ones turn red.")
    steps(doc, [
        "Click D17 first, then hold Shift and click H116 (D17 stays the active cell - check the Name Box shows D17).",
        "Home › Conditional Formatting › New Rule › Use a formula. Type:",
        ("f", "=AND(ISBLANK(D17),COUNTA($D17:$H17)>0)"),
        "Format › Fill › More Colors › Custom 255 / 199 / 206. Font › Bold, colour 156 / 0 / 6. OK, OK.",
        "Click J17, Shift-click J116. New Rule › Use a formula: =LEFT(J17,5)=\"Check\" › same red format.",
        "Same cells, another rule: =J17=\"OK\" › Fill 198 / 239 / 206, font colour 0 / 97 / 0, bold.",
        "Click C17, Shift-click C116. New Rule: =C17=\"Not on list\" › font colour 156 / 0 / 6, bold.",
        "Test: type a Resource on row 17 only - E to H should go red, Status should say 'Check - missing'.",
    ])

    # 7
    part(doc, "Part 7 - Completion check box", "One cell that tells the force whether the form is ready to send.")
    steps(doc, [
        "Pick a spare cell above the table, e.g. I13 (or any empty box). Type the label 'Completion check' in the "
        "cell above it.",
        "In the check cell type:",
        ("f", '=IF(COUNTA(Table1[Resource])=0,"No resources entered yet",IF(COUNTIF(Table1[Status],"Check*")'
              '+(OpName="")+(OpStart="")+(OpEnd="")=0,"READY TO SUBMIT",COUNTIF(Table1[Status],"Check*")'
              '+(OpName="")+(OpStart="")+(OpEnd="")&" item(s) to complete"))'),
        "Conditional formatting on that cell: rule =ISNUMBER(SEARCH(\"to complete\",I13)) › red (as Part 6); rule "
        "=I13=\"READY TO SUBMIT\" › green (as Part 6). Change I13 to the cell you used.",
    ])
    para(doc, "How it counts: each row marked 'Check - missing' is one item; a blank operation name, start date or end "
              "date is one item each. Table1[Status] means the whole Status column, so it keeps working if the table "
              "changes size. If your table has a different name, use it (Table Design › Table Name).")

    # 8
    part(doc, "Part 8 - Guidance tab text",
         "Type each block into column B, one paragraph per cell (Alt+Enter for a new line inside a cell). Merge B:H "
         "for each row and tick Wrap Text. Yellow-fill the heading rows to match the Host Guidance.")
    blocks = [
        ("Title", "EURO 2028 Non-Host Resource Collection Form - Instructions and Guidance"),
        ("Overview:", "Where a national or regional operation (for example Op Brock or Op Stack) runs during the "
                      "EURO 2028 tournament period, the officers it needs are not available for EURO 2028, including "
                      "as mutual aid. This form records the resources your force is committing to that operation, so "
                      "the NPoCC EURO 2028 Police Planning Team can see the total demand on policing and spot pressure "
                      "on tournament resourcing."),
        ("Instructions for forces:", "1. On the Op Data Collection tab, complete the Operation name, Force, From and "
                                     "To dates.\n2. Use one row per resource and rank.\n3. Choose the Resource from "
                                     "the drop-down; the Skill Portfolio fills in automatically.\n4. Choose the Rank, "
                                     "then enter the Number of Officers, Opportunity Cost and Chargeable Hours.\n5. "
                                     "Hours are PER OFFICER PER DAY. Enter 0 if there are none - do not leave a cell "
                                     "blank.\n6. Red cells are missing. The Completion check must read READY TO "
                                     "SUBMIT before you send the form.\n7. Do not add, delete or move rows, columns or "
                                     "tabs. Save as '<Force> - <Operation name>.xlsx'."),
        ("Deadline:", "Returns by [DATE / TIME - TBC] to [EMAIL - TBC]."),
        ("Definitions:", "Number of Officers - officers needed per day. If it changes during the operation, give the "
                         "peak daily figure and explain in Notes.\nOpportunity Cost - normal rostered (tour of duty) "
                         "hours per officer per day; time officers are taken away from their usual duties.\n"
                         "Chargeable Hours - extra (overtime) hours per officer per day caused by the operation. "
                         "Enter 0 if none.\nResource - the national mutual aid skills list. If a skill is not "
                         "listed, choose the closest and explain in Notes."),
        ("Updates:", "If the operation or your commitment changes, send an updated form with the same operation "
                     "name. The latest version will be used."),
        ("Colour key:", "Yellow - choose from the list / required.  White - type here.  Grey - fills in "
                        "automatically.  Red - missing, please complete."),
        ("Before you submit - checklist:", "[ ] Operation name, Force, From and To completed.\n[ ] Every row used "
                                           "has Resource, Rank, Officers, Opportunity Cost and Chargeable Hours.\n"
                                           "[ ] No red cells; Completion check reads READY TO SUBMIT.\n[ ] Saved as "
                                           "'<Force> - <Operation name>.xlsx' and sent by the deadline."),
        ("Contacts:", "NPoCC EURO 2028 Police Planning Team - [NAME / EMAIL / PHONE - TBC]"),
        ("Handling:", "This form and your return are OFFICIAL."),
    ]
    table(doc, ["Heading row", "Text to type underneath"], [[h, t] for h, t in blocks], [4, 13])

    # 9
    part(doc, "Part 9 - Lock, protect and hide")
    steps(doc, [
        "On Op Data Collection, select the input cells: the four operation detail boxes and D17:I116 (Ctrl+click to "
        "add each area).",
        "Right-click › Format Cells › Protection › untick Locked › OK. Everything else stays locked (C and J "
        "formulas, headers, the check box).",
        "Review › Protect Sheet. Tick: Select locked cells, Select unlocked cells, Format columns, Use AutoFilter. "
        "A password is optional - if you set one, record it somewhere safe.",
        "Guidance tab: Review › Protect Sheet (no unlocked cells needed).",
        "Right-click the Lists tab › Hide. Optional: Review › Protect Workbook (Structure) stops tabs being "
        "unhidden, renamed or deleted.",
    ])

    # 10
    part(doc, "Part 10 - Test before sending")
    steps(doc, [
        "Leave the operation boxes empty: the check box should count the missing items.",
        "Pick a Resource on row 17: Skill Portfolio fills in; E to H go red; Status says 'Check - missing'.",
        "Fill the row (use 0 for no overtime): red clears, Status says OK.",
        "Type 25 in an hours cell: it should be refused. Type an end date before the start date: refused.",
        "Fill the operation boxes: the check box reads READY TO SUBMIT in green.",
        "Clear all test entries (select D17:I116 › Delete, and the four boxes) and save.",
    ])

    part(doc, "To confirm with the team")
    bullets(doc, [
        "Chargeable Hours definition for non-host operations - the CFS wording is written for EURO 2028 itself; the "
        "simplified version above just says 'overtime caused by the operation'.",
        "Whether Number of Officers should be the peak daily figure (as written) or per-day rows by date.",
        "Ranks to include (Detective ranks, Police Staff, Specials).",
        "Deadline, return email and contact details.",
    ])
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build())
