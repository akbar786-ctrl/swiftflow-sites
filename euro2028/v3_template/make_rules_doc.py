"""Writes 'EURO 2028 Collection Template - Rules and Formulas Explained' (Word)."""
import json
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output", "EURO_2028_Template_Rules_and_Formulas_Explained.docx")
NAVY = RGBColor(0x1F, 0x38, 0x64)


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def table(doc, head, rows, widths, mono_cols=()):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    for j, h in enumerate(head):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(9), RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, "1F3864")
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            run = cells[j].paragraphs[0].add_run(str(v))
            run.font.size = Pt(8.5 if j in mono_cols else 9)
            if j in mono_cols:
                run.font.name = "Consolas"
    for row in t.rows:
        for j, w in enumerate(widths):
            row.cells[j].width = Cm(w)
    doc.add_paragraph()


def formula(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(text)
    r.font.name, r.font.size = "Consolas", Pt(9)
    shade_p = OxmlElement("w:shd")
    shade_p.set(qn("w:val"), "clear")
    shade_p.set(qn("w:fill"), "F2F2F2")
    p._p.get_or_add_pPr().append(shade_p)


def bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        if isinstance(it, tuple):
            p.add_run(it[0]).bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def para(doc, text, bold=False, italic=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    return p


def build(rows_json, out=OUT):
    R = json.load(open(rows_json))
    md, wt, pl = R["Match Day Resource"], R["Wider Tournament Resource"], R["Pan London"]
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10)
    for n in ("Heading 1", "Heading 2"):
        doc.styles[n].font.name = "Arial"
        doc.styles[n].font.color.rgb = NAVY
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.8)
    for part in (sec.header, sec.footer):
        p = part.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("OFFICIAL")
        r.bold, r.font.color.rgb = True, RGBColor(0x1F, 0x4E, 0x9A)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("EURO 2028 Resource Collection Template\nRules and Formulas Explained")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(18), NAVY
    para(doc, "Covers the Resource Collection Template V2 (General and MPS / Pan London versions): every rule, "
              "condition and formula built in, what it does and why.", italic=True)

    # ------------------------------------------------------------ 1
    doc.add_heading("1. Why the rules exist", level=1)
    para(doc, "Costs are calculated from the hours forces enter, so the template must stop hours being entered "
              "where they do not apply (which would overstate cost) and make it obvious which cells must be "
              "completed. Cells that do not apply to a resource, rank or day are greyed out (hatched) and locked; "
              "everything else is left open.")

    # ------------------------------------------------------------ 2
    doc.add_heading("2. Column layout the rules refer to", level=1)
    para(doc, "Every day (Match Day tab) or risk level (other tabs) has three columns, always in this order: "
              "1 Total Officers required, 2 Opportunity Cost (hours per officer), 3 Chargeable Hours (hours per "
              "officer).")
    table(doc, ["Tab", "Risk level", "Columns"], [
        ["Match Day", "Low", "H:S  (MD-2 H:J, MD-1 K:M, MD N:P, MD+1 Q:S)"],
        ["Match Day", "Medium", "T:AE (MD-2 T:V, MD-1 W:Y, MD Z:AB, MD+1 AC:AE)"],
        ["Match Day", "High", "AF:AQ (MD-2 AF:AH, MD-1 AI:AK, MD AL:AN, MD+1 AO:AQ)"],
        ["Match Day", "Super High", "AR:BC (MD-2 AR:AT, MD-1 AU:AW, MD AX:AZ, MD+1 BA:BC)"],
        ["Wider Tournament", "Low / Medium / High / Super High", "G:I / J:L / M:O / P:R"],
        ["Pan London", "Low / Medium / High / Super High", "F:H / I:K / L:N / O:Q"],
    ], [3.2, 4.3, 9.5])
    para(doc, "So on the Match Day tab: Opportunity Cost columns are I, L, O, R, U, X, AA, AD, AG, AJ, AM, AP, AS, "
              "AV, AY, BB and Chargeable columns are J, M, P, S, V, Y, AB, AE, AH, AK, AN, AQ, AT, AW, AZ, BC.")

    # ------------------------------------------------------------ 3
    doc.add_heading("3. Fixed hatching rules (pre-filled lines)", level=1)
    para(doc, "Each pre-filled line is identified by its Deployment Location, Resource, Resource Origin and Rank "
              "columns. The rules below are applied in this order; a cell is hatched if any rule applies.")
    table(doc, ["#", "Rule", "Cells hatched", "Why"], [
        ["1", "Match Day only lines - Resource contains '(Match Day' / '(Match Day Only)', plus location-based "
              "public order lines (In Stadium, Stadium Vicinity, Town Centre)",
         "All three columns of Match Day -2, Match Day -1 and Match Day +1",
         "The resource is only deployed on the match day itself."],
        ["2", "Non-Match Day lines - Resource contains 'Non-Match Day'", "All three Match Day columns",
         "The resource is only deployed on the days around the match."],
        ["3", "PSU / BDU number lines - Rank is 'Level 2 PSU' or 'Level 3 BDU'",
         "Opportunity Cost and Chargeable (every day / risk level)",
         "These lines hold the NUMBER of units. Hours are entered per officer on the PSU / Serial "
         "configuration lines; hours here would be double counted."],
        ["4", "In Stadium lines (any rank, either origin)", "Chargeable", "Agreed costing rule: In Stadium "
                                                                          "resources are costed on Opportunity Cost."],
        ["5", "Mutual Aid lines (not In Stadium)", "Opportunity Cost", "Agreed costing rule: only Chargeable "
                                                                       "hours are costed for mutual aid."],
        ["6", "Force Supplied Inspector or Detective Inspector (not In Stadium)", "Chargeable",
         "Agreed costing rule: Force Supplied Inspectors are not costed on chargeable hours."],
    ], [0.7, 5.5, 4.3, 6.5])
    doc.add_heading("Rows each rule applies to", level=2)
    table(doc, ["Rule", "Match Day rows", "Wider Tournament rows", "Pan London rows"], [
        ["1 Match Day only", md["mdonly"], "-", "-"],
        ["2 Non-Match Day", md["nonmd"], "-", "-"],
        ["3 PSU / BDU number", md["unit"], wt["unit"], pl["unit"]],
        ["4 In Stadium", md["stadium"], "-", "-"],
        ["5 Mutual Aid", md["ma"], wt["ma"], pl["ma"]],
        ["6 FS Inspector", md["fsinsp"], wt["fsinsp"], pl["fsinsp"]],
    ], [3.0, 7.0, 3.6, 3.4])
    doc.add_heading("How a hatched cell is built", level=2)
    bullets(doc, [
        ("Appearance: ", "Fill > Pattern Style '25% Gray' (lightGray), pattern colour grey 808080, background light "
                         "grey D9D9D9."),
        ("Locked: ", "Format Cells > Protection > Locked ticked (takes effect once the sheet is protected)."),
        ("Typing blocked: ", "Data > Data Validation > Allow: Custom, formula below, error alert 'Not required - "
                             "This cell is not required for this resource, rank or day. Please leave it blank.'"),
    ])
    formula(doc, "=FALSE")
    para(doc, "Explanation: a Custom validation accepts an entry only when its formula is TRUE. FALSE is never TRUE, "
              "so any typed value is rejected and the error message is shown. 'Ignore blank' is ticked, so an empty "
              "cell is still valid. (Validation does not stop pasting - sheet protection does.)")

    # ------------------------------------------------------------ 4
    doc.add_heading("4. Additional Resources lines (rules that react to the origin and rank chosen)", level=1)
    para(doc, "These lines are blank until a force fills them in, so the hatching cannot be fixed in advance. "
              "Conditional formatting greys the cell automatically and validation blocks entry once the Resource "
              "Origin and Rank are chosen. Rows: Match Day 347-371, Wider Tournament 93-102, Pan London 37-46.")
    doc.add_heading("4a. Mutual Aid - Opportunity Cost greyed out", level=2)
    para(doc, "Applies to the Opportunity Cost columns of the Additional Resources rows. Conditional formatting "
              "(format = the grey hatch, 'Stop If True' ticked):")
    formula(doc, 'Match Day:         =LEFT($F347,6)="Mutual"')
    formula(doc, 'Wider Tournament:  =LEFT($E93,6)="Mutual"')
    formula(doc, 'Pan London:        =LEFT($D37,6)="Mutual"')
    bullets(doc, [
        ("$F347 ", "is the Resource Origin of the same row (column F on Match Day, E on Wider Tournament, D on Pan "
                   "London). The $ fixes the column so every hours column looks at the origin; the row number has no "
                   "$ so each row looks at its own origin."),
        ('LEFT(...,6)="Mutual" ', "is TRUE when the origin starts with 'Mutual' - this catches 'Mutual Aid', "
                                  "'Mutual Aid Supplied' and 'Mutual Aid Requested' alike."),
    ])
    para(doc, "Validation on the same cells (Custom) - blocks typing when the rule is TRUE:")
    formula(doc, 'Match Day:         =NOT(LEFT($F347,6)="Mutual")')
    para(doc, "NOT(...) turns the rule round: entry is allowed only when the line is NOT mutual aid. "
              "(Wider Tournament / Pan London use $E93 / $D37.)")
    doc.add_heading("4b. Force Supplied Inspector - Chargeable greyed out", level=2)
    para(doc, "Applies to the Chargeable columns of the Additional Resources rows. Conditional formatting (grey "
              "hatch, 'Stop If True'):")
    formula(doc, 'Match Day:  =AND(LEFT($F347,5)="Force",OR($G347="Inspector",$G347="Detective Inspector"))')
    formula(doc, 'WT:         =AND(LEFT($E93,5)="Force",OR($F93="Inspector",$F93="Detective Inspector"))')
    formula(doc, 'Pan London: =AND(LEFT($D37,5)="Force",OR($E37="Inspector",$E37="Detective Inspector"))')
    bullets(doc, [
        ('LEFT($F347,5)="Force" ', "- the origin starts with 'Force' (Force Supplied)."),
        ("$G347 ", "- the Rank of the same row (G on Match Day, F on Wider Tournament, E on Pan London)."),
        ("OR(...) ", "- the rank is Inspector or Detective Inspector."),
        ("AND(...) ", "- both must be true: a Force Supplied Inspector. Chief Inspector and above are not affected."),
    ])
    para(doc, "Validation on the same cells (Custom):")
    formula(doc, 'Match Day:  =NOT(AND(LEFT($F347,5)="Force",OR($G347="Inspector",$G347="Detective Inspector")))')

    # ------------------------------------------------------------ 5
    doc.add_heading("5. Locking", level=1)
    bullets(doc, [
        ("Unlocked (forces can type): ", "every non-hatched number cell on the three data tabs, plus the Resource, "
                                         "Origin and Rank cells of the Additional Resources rows (Match Day E347:G371, "
                                         "Wider Tournament D93:F102, Pan London C37:E46)."),
        ("Locked: ", "hatched cells, labels, headers and totals."),
        ("Switching it on: ", "Review > Protect Sheet on each tab. Until then the validation in sections 3 and 4 "
                              "still blocks typing in hatched cells."),
    ])

    # ------------------------------------------------------------ 6
    doc.add_heading("6. Colours", level=1)
    table(doc, ["Item", "Colour (hex)"], [
        ["Match Day risk headers (row 3/4)", "Low 2E7D32, Medium D4A017, High BF5B04, Super High 8B0000"],
        ["Totals headers BD4 / BG4 / BJ4 / BM4", "Low 2E7D32, Medium D4A017, High BF5B04, Super High 8B0000"],
        ["High cell shading (3 measures)", "FEF4EA, FDEBDA, FCE2CA (replaced the previous pink)"],
        ["Super High cell shading (3 measures)", "FFF3F3, FFE8E8, FFDDDD"],
        ["Wider Tournament / Pan London row 3 headers", "Low C8E6C9, Medium F0DC8B, High F8C471, Super High EF9A9A"],
        ["Wider Tournament / Pan London Super High cells", "FFE8E8 (replaced the previous blue)"],
        ["Hatched (not required)", "25% Gray pattern, grey 808080 on D9D9D9"],
    ], [7, 10])

    # ------------------------------------------------------------ 7
    doc.add_heading("7. Other changes", level=1)
    bullets(doc, [
        "Header text corrected: 'requered' to 'required', 'Catagories' to 'Categories'; totals headers relabelled "
        "Low / Medium / High / Super High (previously 'Match Day -6 / -7 / +5').",
        "Guidance tab completed (two versions: General, and MPS with the Pan London instruction and note).",
        "General version: Pan London tab removed.",
    ])
    doc.add_heading("Unchanged from the original template", level=2)
    bullets(doc, [
        "Conditional formatting 'Cell value equal to 0' rules on headers and data (formatting of zeros).",
        "Drop-down lists and input prompts in the label columns (e.g. Additional Resources origin / rank, Data Info "
        "lists).",
        "Automatic Total Figures formulas (columns BD:BO on Match Day).",
    ])

    # ------------------------------------------------------------ A
    doc.add_heading("Appendix A - Missing-hours highlight (removed; to add manually)", level=1)
    para(doc, "Turns an hours cell red when officers have been entered for that day but the hours cell is blank. "
              "Apply it ONLY to the open (white) Opportunity Cost and Chargeable cells - never to hatched cells, "
              "because they are always blank and would turn red.")
    para(doc, "Opportunity Cost columns (officers are one column to the left). Example with the first selected cell "
              "I8 on Match Day:", bold=True)
    formula(doc, "=AND(N(H8)>0,ISBLANK(I8))")
    para(doc, "Chargeable columns (officers are two columns to the left). Example with the first selected cell J8:",
         bold=True)
    formula(doc, "=AND(N(H8)>0,ISBLANK(J8))")
    bullets(doc, [
        ("N(H8)>0 ", "- the officers cell holds a number greater than 0. N() turns text or blank into 0 so the rule "
                     "cannot error."),
        ("ISBLANK(I8) ", "- the hours cell is empty (0 counts as completed)."),
        ("AND(...) ", "- both true: officers entered but hours missing. Format: light red fill FFC7CE, dark red bold "
                      "font 9C0006."),
        ("Relative references: ", "write the formula for the active (first) cell of the selection; Excel adjusts it "
                                  "for every other cell. Place this rule BELOW the Additional Resources hatch rules in "
                                  "Manage Rules so the hatching shows first."),
    ])

    # ------------------------------------------------------------ B
    doc.add_heading("Appendix B - Helper-formula method (for rebuilding the hatching by hand)", level=1)
    para(doc, "An alternative that gives the same hatching with a handful of formulas. Hidden helper columns work "
              "out what each line is; one conditional formatting rule and one validation rule then hatch and "
              "block the right cells. Shown for Match Day; Wider Tournament and Pan London are the same idea "
              "without the day rule.")
    table(doc, ["Cell", "Formula", "What it does"], [
        ["BQ8", '=IF(F8<>"",F8,BQ7)', "Resource Origin of the line. Merged cells only hold text in their first row, "
                                       "so this carries the last origin down."],
        ["BR8", '=IF(D8<>"",D8,BR7)', "Deployment Location, carried down the same way."],
        ["BS8", '=IF(E8<>"",SUBSTITUTE(E8," ",""),BS7)', "Resource name with spaces removed, so 'Non-Match Day' "
                                                         "and 'Non- Match Day' read the same."],
        ["BT8", '=IF(ISNUMBER(SEARCH("Non-Match",BS8)),"NONMD",IF(OR(ISNUMBER(SEARCH("(MatchDay",BS8)),'
                'AND(ROW()<=48,OR(BR8="In Stadium",BR8="Stadium Vicinity",BR8="Town Centre"))),"MDONLY","ALL"))',
         "Day rule: NONMD, MDONLY or ALL."],
        ["BU8", '=IF(OR(TRIM(G8)="",TRIM(G8)="Rank"),"",IF(OR(TRIM(G8)="Level 2 PSU",TRIM(G8)="Level 3 BDU"),'
                '"UNIT",IF(BR8="In Stadium","STADIUM",IF(LEFT(BQ8,6)="Mutual","MA",IF(OR(TRIM(G8)="Inspector",'
                'TRIM(G8)="Detective Inspector"),"FSINSP","OPEN")))))',
         "Hours rule, checked in priority order: UNIT, STADIUM, MA, FSINSP, otherwise OPEN. Blank for "
         "headings."],
        ["H373", "=INT(MOD(COLUMN()-8,12)/3)", "Day number of the column: 0 MD-2, 1 MD-1, 2 MD, 3 MD+1."],
        ["H374", "=MOD(COLUMN()-8,3)", "Measure of the column: 0 Officers, 1 Opportunity Cost, 2 Chargeable."],
    ], [1.4, 9.0, 6.6], mono_cols=(1,))
    para(doc, "Hatch rule (conditional formatting on H8:BC371, H8 as the active cell):", bold=True)
    formula(doc, '=AND($BU8<>"",OR(AND($BT8="MDONLY",H$373<>2),AND($BT8="NONMD",H$373=2),AND($BU8="UNIT",H$374>0),'
                 'AND($BU8="STADIUM",H$374=2),AND($BU8="MA",H$374=1),AND($BU8="FSINSP",H$374=2)))')
    para(doc, "Read it as: the line is a data line, AND one of - Match Day only and not the Match Day column; "
              "Non-Match Day and the Match Day column; a PSU number line and an hours column; In Stadium and "
              "Chargeable; Mutual Aid and Opportunity Cost; Force Supplied Inspector and Chargeable. The validation "
              "uses the same formula wrapped in NOT(...). This method was checked against the template and gives "
              "identical hatching.")

    # ------------------------------------------------------------ C
    doc.add_heading("Appendix C - Points to confirm", level=1)
    bullets(doc, [
        "Detective Inspector is treated the same as Inspector; Chief Inspector and above are left open.",
        "Mutual Aid is taken from the Resource Origin column (F on Match Day). Rows 33-40 show Force Supplied in F "
        "while the old hidden column B says Mutual Aid - F is used.",
        "Wider Tournament protest removal lines (rows 45-47, 54-56) have no rank.",
    ])
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else OUT))
