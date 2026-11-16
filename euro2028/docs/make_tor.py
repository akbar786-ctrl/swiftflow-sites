"""Builds the Terms of Reference (Word) for the EURO 2028 resource collection and Power BI dashboard."""
import datetime as dt
import os
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "EURO_2028_Resource_Dashboard_Terms_of_Reference.docx")
NAVY = RGBColor(0x1F, 0x38, 0x64)
STATUS = {"Supported": "C6EFCE", "Supported once reference data is added": "FFEB9C",
          "Needs new data or a business rule": "F8CBAD"}


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def table(doc, head, rows, widths=None, status_col=None, font=9):
    t = doc.add_table(rows=1, cols=len(head))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(head):
        c = t.rows[0].cells[j]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold, r.font.size, r.font.color.rgb = True, Pt(font), RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, "1F3864")
    for row in rows:
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            run = cells[j].paragraphs[0].add_run(str(v))
            run.font.size = Pt(font)
            if status_col is not None and j == status_col and v in STATUS:
                shade(cells[j], STATUS[v])
    if widths:
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
    doc.add_paragraph()
    return t


def bullets(doc, items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0])
            r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def para(doc, text, bold=False, italic=False, size=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    if size:
        r.font.size = Pt(size)
    return p


def build(out=OUT):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10.5)
    for name in ("Heading 1", "Heading 2", "Heading 3"):
        doc.styles[name].font.name = "Arial"
        doc.styles[name].font.color.rgb = NAVY
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2)
    for part in (sec.header, sec.footer):
        p = part.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("OFFICIAL")
        r.bold, r.font.color.rgb = True, RGBColor(0x1F, 0x4E, 0x9A)

    # ------------------------------------------------------------------ title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("EURO 2028 Policing\nResource Forecast Collection and Power BI Dashboard")
    r.bold, r.font.size, r.font.color.rgb = True, Pt(20), NAVY
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Terms of Reference")
    r.bold, r.font.size = True, Pt(16)
    table(doc, ["Item", "Detail"], [
        ["Document", "Terms of Reference - EURO 2028 Resource Forecast Collection and Dashboard"],
        ["Version", "0.1 DRAFT for review"],
        ["Date", dt.date.today().strftime("%d %B %Y")],
        ["Owner", "NPoCC Euro 2028 Police Planning Team - [named lead TBC]"],
        ["Classification", "OFFICIAL"],
        ["Related documents", "Euro_2028_Question_Prompts KL; Resource Forecast Collection Form (v3); "
                              "Power BI build guide; Legacy collection workbook audit"],
    ], widths=[4, 12.5])

    # ------------------------------------------------------------- 1 purpose
    doc.add_heading("1. Purpose", level=1)
    para(doc, "These Terms of Reference set out the purpose, scope, deliverables, methods, responsibilities and "
              "governance for the EURO 2028 Resource Forecast Collection and the Power BI dashboard built from it. "
              "They follow the structure of the Euro 2028 Question Prompts document (Context, Schedule, Resourcing, "
              "Budgeting, Other, Existing Dashboards and Models) and show, question by question, how each will be "
              "answered and what is still needed.")

    # ------------------------------------------------------------- 2 context
    doc.add_heading("2. Context", level=1)
    para(doc, "The dashboard will be used to support capability and capacity planning, specifically for Euro 2028 "
              "policing, both at national and local level. It will be used by multiple client groups:")
    bullets(doc, ["NPoCC Euro 2028 Police Planning Team", "NPoCC Ops", "Strategic Planning and Performance",
                  "NPCC portfolio leads", "Host Force planning teams"])
    para(doc, "Background", bold=True)
    bullets(doc, [
        "A first round of Euro 2028 data collection used an Excel workbook in which forces entered requirements and "
        "a set of hidden background sheets (Demand Sheet, Cost Sheet, Cost Calculator, Fleet and others) calculated "
        "demand and cost. Two existing products were built from that round (section 8).",
        "An audit of that workbook found the calculation logic hard-coded into tens of thousands of cells, with "
        "broken references (#REF!), copy errors, inconsistent rules and a calendar still set to 2025. The figures "
        "were difficult to check, maintain or combine across forces.",
        "The project replaces this with a simple collection form completed by forces, and moves all processing "
        "and calculation into Power BI (Power Query and DAX), where the rules are written once, visible and "
        "testable.",
    ])

    # ------------------------------------------------------------ 3 objectives
    doc.add_heading("3. Objectives", level=1)
    bullets(doc, [
        "Collect consistent resource forecasts from every force using one straightforward form.",
        "Combine all force returns automatically into a single national dataset.",
        "Reproduce the agreed costing and demand rules in Power BI, transparently and repeatably.",
        "Answer the Schedule, Resourcing, Budgeting and Other questions in the Question Prompts document at force, "
        "regional (where data allows) and national level.",
        "Report data-quality issues so returns can be corrected before figures are used.",
        "Keep the solution maintainable by the NPoCC team without specialist support.",
    ])

    # --------------------------------------------------------------- 4 scope
    doc.add_heading("4. Scope", level=1)
    para(doc, "In scope", bold=True)
    bullets(doc, [
        "The Resource Forecast Collection Form: two sheets completed by forces - Match Day Resource and Wider "
        "Tournament Resource - plus Guidance.",
        "Power BI: loading all returns, reshaping them into master tables, reference tables (rates, settings, "
        "Wider Tournament days, fixture schedule, skillset mapping), DAX measures and report pages.",
        "Cost and resource breakdowns by force, skillset, section, resource, rank, origin (force supplied / mutual "
        "aid), risk level, relative day and date.",
        "Data-quality reporting, validation against the legacy workbook, and user and build guidance.",
    ])
    para(doc, "Out of scope (unless agreed through change control)", bold=True)
    bullets(doc, [
        "Changing the legacy collection workbook or the first-round dashboards.",
        "Operational deployment decisions - the dashboard informs planning; it does not allocate resources.",
        "Collecting new data types (e.g. equipment, vehicles, horses, force capacity baselines) until the business "
        "rules and sources in section 11 are agreed.",
    ])

    # ---------------------------------------------------------- 5 deliverables
    doc.add_heading("5. Deliverables", level=1)
    table(doc, ["#", "Deliverable", "Description", "Status"], [
        ["D1", "Collection Form (v3)", "One row per resource / origin / rank. Match Day: Low, Medium, High, Super "
         "High x Match Day -2, -1, Match Day, +1 x Officers, Normal Tour of Duty, Extra Hours. Wider Tournament: the "
         "same per risk level for a typical tournament day. Validation on entry, hatched cells where not required, "
         "hidden one-row header for Power BI.", "Built - for review"],
        ["D2", "Power BI data model", "Power Query reads every return from a folder and builds master tables "
         "(Match Day, Wider Tournament, Fixture Calendar), dimension tables and a Data Quality table.",
         "Designed and tested on the previous layout; to be rebuilt step by step for v3"],
        ["D3", "Reference tables", "Hourly rates by rank and origin, settings, Wider Tournament days per risk level, "
         "fixture schedule, Mercury skillset mapping.", "Rates and days in place; schedule and skillset mapping "
         "required"],
        ["D4", "DAX measures", "Officers, officer deployments, hours, overtime cost, normal duty cost, mutual aid "
         "cost, totals by any dimension.", "Drafted"],
        ["D5", "Report pages", "Overview, Schedule, Match Day, Fixture Calendar, Wider Tournament, Budget by "
         "portfolio, Resource breakdown, Data Quality.", "Mock-ups produced"],
        ["D6", "Validation pack", "Tests against legacy outputs and synthetic test returns with expected results.",
         "Legacy logic reproduced with 0 differences"],
        ["D7", "Guidance", "Force completion guidance; step-by-step Power BI build and refresh guide.", "Drafted"],
    ], widths=[1, 3.5, 9, 3.5])

    # ------------------------------------------- 6 question-by-question
    doc.add_heading("6. Dashboard questions and how they will be answered", level=1)
    para(doc, "Status key: Supported = answered from the forms and existing reference data; Supported once "
              "reference data is added = the model handles it, a reference table must be supplied; Needs new data "
              "or a business rule = cannot be answered until the project team provides a source or rule.",
         italic=True, size=9)
    H = ["Ref", "Question (from Question Prompts)", "How the dashboard answers it", "Status"]
    W = [1.1, 6.2, 7.2, 2.5]
    S, A, N = "Supported", "Supported once reference data is added", "Needs new data or a business rule"
    doc.add_heading("6.1 Schedule", level=2)
    table(doc, H, [
        ["S1", "For each venue, when are the fixtures being played?", "Fixture schedule table (venue, date) shown "
         "as a calendar and list.", A],
        ["S2", "For each venue, what time is the fixture being played?", "Kick-off time held in the fixture "
         "schedule table.", A],
        ["S3", "For each venue, which of competing nations are playing? (post qualifying draw)", "Team fields in the "
         "schedule table, completed after the draw.", A],
        ["S4", "For each fixture, when is MD-2, MD-1, and MD+1?", "Calculated automatically from the fixture date.", S],
        ["S5", "For each venue, which other venues are playing on the same day and at what times?", "Schedule page "
         "filtered by date, listing all venues and kick-off times.", A],
        ["S6", "How many fixtures are played at each venue for a specific date range? (calendar view)", "Calendar "
         "view with a date-range slicer.", A],
    ], widths=W, status_col=3)
    doc.add_heading("6.2 Resourcing", level=2)
    table(doc, H, [
        ["R1", "For each Mercury defined skillset, how many wider tournament resources are deployed on a low, "
         "medium, and high-risk scenario?", "Wider Tournament master table by risk level (Super High also "
         "available), grouped by Mercury skillset via the skillset mapping.", A],
        ["R2-R5", "In accordance with the schedule and for each Mercury defined skillset, how many resources are "
         "deployed on MD-2 / MD-1 / MD / MD+1 in a low, medium, and high-risk scenario?", "Match Day master table "
         "by relative day and risk level, grouped by skillset; placed on dates via the fixture schedule.", A],
        ["R6", "For a defined date range, how many resources (categorised by Mercury skillset) are deployed each day "
         "within the range (both matchday and wider tournament)?", "Matchday part: Fixture Calendar table. Wider "
         "tournament part needs a rule for which risk level applies on each tournament date (the forms give a "
         "typical day per risk level; the national assumption gives day counts, not dates).", N],
        ["R7", "Categorised by Mercury skillset, how many resources are deployed for each day of tournament in "
         "accordance with the schedule and risk scenario specific to each host force?", "Fixture Calendar with the "
         "risk level of each fixture for each host force in the schedule table.", A],
        ["R8", "In accordance with the above scenarios, the dashboard will provide updated force, regional & "
         "national capacity against defined portfolios / skills.", "Requires a capacity baseline (available "
         "officers by skillset per force) and a force-to-region mapping; demand can then be compared with "
         "capacity.", N],
    ], widths=W, status_col=3)
    doc.add_heading("6.3 Budgeting", level=2)
    table(doc, H, [
        ["B1", "For each Mercury defined skillset, what is the cost to deploy the proposed amount of resources, both "
         "at force and national level?", "Cost measures grouped by skillset, filtered by force or national.", A],
        ["B2", "For resources deployed to a fan zone, what is the cost ...?", "Deployment Location = Fan Zone "
         "(Wider Tournament) and the Fan Zone sections.", S],
        ["B3", "For resources deployed to night-time economy (community policing), what is the cost ...?",
         "Deployment Location = Wider Community, if this is confirmed as the night-time economy.", A],
        ["B4", "For armed policing, what is the cost ...?", "Sections Armed Policing, Additional Armed Policing "
         "(Armed Response Overlay) and Fan Zone Armed Policing.", S],
        ["B5", "For roads policing, what is the cost ...?", "Section Roads Policing.", S],
        ["B6", "For drone policing, what is the cost ...?", "Section Unmanned Aerial Systems (Drones).", S],
        ["B7", "For CBRN policing, what is the cost ...?", "Section CBRN.", S],
        ["B8", "For mutual aid supplied resources, what is the cost ...?", "Resource Origin = Mutual Aid Supplied "
         "(subject to the mutual aid costing decision, section 11).", S],
        ["B9", "For force supplied resources, what is the cost ...?", "Resource Origin = Force Supplied.", S],
    ], widths=W, status_col=3)
    doc.add_heading("6.4 Other", level=2)
    table(doc, H, [
        ["O1", "Interpreting the proposed level 2 public order resources, how many protected carriers are required "
         "both at a local and national level?", "The form collects Public Order Protective Carrier Drivers; a rule "
         "converting drivers (or PSUs) to carriers is required.", A],
        ["O2", "Interpreting the proposed mounted officer resources, how many actual horses are required ...?",
         "Needs a horses-per-officer rule from the Mounted portfolio lead.", N],
        ["O3", "Interpreting the proposed counter drone officer resources, how many counter drone equipment are "
         "required ...?", "Needs an equipment-per-officer (or per team) rule from the portfolio lead.", N],
        ["O4", "Interpreting the proposed live facial recognition officer resources, how many LFR Vans are "
         "required ...?", "Needs a vans-per-officer (or per team) rule.", N],
        ["O5", "Should Op Brock be called, how many additional resources (categorised by Mercury skillset) will be "
         "required to police the operation?", "Not collected on the current form. Needs either an extra scenario "
         "on the form or an uplift rule.", N],
        ["O6", "Should Op Stack be called, how many additional resources ... will be required?", "As O5.", N],
        ["O7", "Should Israel qualify, how many additional resources ... will be required?", "As O5; could be "
         "modelled as a scenario uplift on affected fixtures once the draw is known.", N],
        ["O8", "Categorised by Mercury skillset, how many resources are required for the national policing reserve "
         "and what is the national impact of this?", "Needs the reserve definition (number or percentage by "
         "skillset) and capacity baselines (see R8).", N],
    ], widths=W, status_col=3)

    # ------------------------------------------------------ 7 methodology
    doc.add_heading("7. Data collection and calculation method", level=1)
    doc.add_heading("7.1 The collection form", level=2)
    bullets(doc, [
        ("Two sheets completed by forces: ", "Match Day Resource and Wider Tournament Resource. Every row is one "
         "resource, from one origin (Force Supplied or Mutual Aid Supplied), at one rank, with Overlay and "
         "Deployment Location where relevant (In Stadium / Stadium Vicinity / Any other Location; Wider Community "
         "/ Fan Zone)."),
        ("Risk levels: ", "Low, Medium, High and Super High. Forces know their fixture dates but not the competing "
         "nations, so they complete every risk level; the planning team applies the appropriate level per fixture."),
        ("Match Day measures: ", "for each of Match Day -2, -1, Match Day and +1: Total Officers required, Normal "
         "Tour of Duty hours per officer and Extra Hours / Over Time per officer."),
        ("Role-based resources ", "(Public Order roles, CBRN, Armed Response Overlay) are one row per role with the "
         "agreed rank, e.g. Bronze Commander - Inspector; Tactical Commander (Silver) - Chief Inspector; Specialist "
         "Strategic Firearms Commander - Assistant Chief Constable."),
        ("Validation: ", "whole numbers for officers, 0-24 hours, hatched/locked cells where an entry is not "
         "required (mutual aid Normal Tour of Duty), Additional Resources rows with origin and rank lists."),
        ("Returns: ", "saved as '<Force name>.xlsx' in one folder; rows and columns must not be inserted or moved."),
    ])
    doc.add_heading("7.2 Processing in Power BI", level=2)
    bullets(doc, [
        "Power Query reads only the two collection sheets of every return, uses the hidden header row, and turns "
        "the grid into master tables (one row per force, resource, origin, rank, risk level and day).",
        "Reference tables hold everything that is not collected from forces: rates, settings, Wider Tournament "
        "days, fixture schedule and the Mercury skillset mapping. They are maintained by the planning team in one "
        "place.",
        "DAX measures produce the totals shown on the report pages; every figure can be filtered by force, "
        "skillset, section, resource, rank, origin, risk level, day and date.",
    ])
    doc.add_heading("7.3 Calculation rules (as agreed with the project team)", level=2)
    table(doc, ["Rule", "Agreed position", "Source"], [
        ["Officer deployments", "Officers x days (an officer on duty for three days counts three).", "Design"],
        ["Hourly rates", "Pulled through by rank; mutual aid rates differ from force supplied rates. June 2028 "
         "rates from the original workbook's Baseline Costs.", "Project team answers 4 and 7; 'Mutual aid costs are "
         "different'"],
        ["Force supplied Inspectors and above", "Costed at GBP 0.", "Project team answer 5"],
        ["Overtime cost", "Officers x Extra Hours x hourly rate.", "Original workbook"],
        ["Normal Tour of Duty", "Costed for all officers (Officers x Normal ToD hours x rate) and shown separately "
         "so the original overtime-only basis can still be compared.", "Project team answer 3"],
        ["Mutual aid PSU package", "GBP 35,716 per PSU covers the travel day in, the match day and the travel day "
         "out. The v3 form collects mutual aid as officers by rank - costing method to be confirmed (section 11).",
         "Project team answer 6"],
        ["Wider Tournament", "Typical-day figure x number of days at each risk level: Low 7, Medium 11, High 11, "
         "Super High 3. Super High retained.", "National assumptions; answer 8"],
        ["Fixture calendar", "MD-2, MD-1, MD+1 placed around each fixture date. Where one force has two "
         "fixture-days on the same date at the same risk level, one day is used (Match Day > MD-1 > MD-2 > MD+1), "
         "as in the original workbook.", "Original workbook - to confirm"],
        ["Fixtures", "Not yet known; the model is built to accept them when confirmed.", "Project team"],
        ["Removed resources", "Items removed from the new template are not carried forward.", "Answer 9"],
        ["Other pages", "Other first-round pages are brought into Power BI in case they are needed later.",
         "Answer 10"],
    ], widths=[3.5, 9.5, 4])

    # --------------------------------------------------------- 8 existing
    doc.add_heading("8. Existing dashboards and models", level=1)
    table(doc, ["Product", "Description", "Relationship to this project"], [
        ["1) NPoCC Ops Dashboard - Summer Demand 2028 Resource Dashboard.xlsx", "Designed after the first round of "
         "Euro 2028 data collection. Filters by force, skill, tier and portfolio; daily demand by skillset; "
         "a tab per Mercury defined skillset (e.g. POPS).", "The Power BI dashboard will provide the same views "
         "from the new data and should use the same skillset names, tiers and portfolios, so users recognise them."],
        ["2) NPoCC Euro 2028 Police Planning Team Model - FOR_DEMO_EU28_COLLECTION_MODEL.xlsx", "Designed after the "
         "first round of Euro 2028 data collection.", "To be reviewed so its views are reproduced or consolidated "
         "in the new dashboard."],
    ], widths=[5, 6, 6])

    # --------------------------------------------------- 9 data quality
    doc.add_heading("9. Data quality and validation", level=1)
    bullets(doc, [
        ("Data Quality page: ", "lists returns that are not on the current form, numbers with no rank or resource, "
         "hours outside 0-24, hours with no officers, missing rates, and fixture forces with no return."),
        ("Legacy reproduction: ", "the new calculation logic was run on a force's completed first-round return and "
         "matched the original workbook's daily officers, daily cost and Wider Tournament total with 0 differences "
         "(also with figures entered as Super High). Known defects in the original workbook were identified and "
         "explained in the audit rather than reproduced silently."),
        ("Test pack: ", "synthetic returns with expected results, used to prove each Power BI build before real data "
         "is loaded."),
        ("Principles: ", "no formula is assumed correct because it existed in the old workbook; no business rule or "
         "cost assumption is invented - all are listed in section 11 until confirmed."),
    ])

    # ---------------------------------------------- 10 roles
    doc.add_heading("10. Roles and responsibilities", level=1)
    table(doc, ["Role", "Responsibility", "Name"], [
        ["Senior responsible owner", "Approves the ToR, scope changes and outputs.", "[TBC]"],
        ["Project lead - NPoCC Euro 2028 Police Planning Team", "Day-to-day ownership; maintains reference tables; "
         "chases forces; signs off figures.", "[TBC]"],
        ["Power BI developer / analyst", "Builds and maintains the data model, measures and report pages; runs "
         "validation.", "[TBC]"],
        ["NPCC portfolio leads", "Confirm skillset mapping and conversion rules (carriers, horses, counter-drone "
         "equipment, LFR vans), reserve requirements.", "[TBC]"],
        ["Strategic Planning and Performance", "Assurance of method and outputs.", "[TBC]"],
        ["Host Force planning teams", "Complete and return the collection form; correct data-quality issues.",
         "[TBC]"],
        ["NPoCC Ops", "User of operational views.", "[TBC]"],
    ], widths=[5, 9, 3])

    # ---------------------------------------------- 11 decisions
    doc.add_heading("11. Decisions and information required", level=1)
    table(doc, ["#", "Decision / information", "Needed for", "Owner"], [
        ["1", "Mercury skillset mapping: each form resource / rank to its Mercury defined skillset (and tier / "
         "portfolio).", "R1-R7, B1, O5-O8", "[TBC]"],
        ["2", "Fixture schedule: venue, date, kick-off time, host force(s), risk level per fixture; teams after the "
         "qualifying draw.", "S1-S6, R2-R7", "[TBC]"],
        ["3", "Which risk level applies on each Wider Tournament date (or confirm WT is reported as totals only).",
         "R6, R7", "[TBC]"],
        ["4", "Mutual aid costing in v3: officer hourly rate x hours, or the GBP 35,716 package per PSU (and how "
         "officers convert to PSUs).", "B8, totals", "[TBC]"],
        ["5", "Normal Tour of Duty for mutual aid: confirm it is not collected (hatched) - so mutual aid normal "
         "duty cost is nil - given Normal ToD is costed for all officers.", "Cost totals", "[TBC]"],
        ["6", "Night-time economy = Wider Community deployment location?", "B3", "[TBC]"],
        ["7", "Conversion rules: protected carriers per driver/PSU, horses per mounted officer, counter-drone "
         "equipment, LFR vans.", "O1-O4", "[TBC]"],
        ["8", "Op Brock, Op Stack and Israel-qualification scenarios: extra collection or uplift rules.", "O5-O7",
         "[TBC]"],
        ["9", "National policing reserve definition and force capacity baselines; force-to-region mapping.",
         "R8, O8", "[TBC]"],
        ["10", "Fixture-calendar rule when fixture-days overlap (one day per date, as the original) - confirm.",
         "Calendar", "[TBC]"],
        ["11", "Ranks offered for Mounted, Dogs, Roads Policing and similar (Inspector / Sergeant / Constable).",
         "Form", "[TBC]"],
    ], widths=[0.8, 9.2, 3.5, 3.5])

    # ---------------------------------------------- 12 assumptions & risks
    doc.add_heading("12. Assumptions, constraints and dependencies", level=1)
    bullets(doc, [
        "Forces return the form unchanged in structure; the file name identifies the force.",
        "Rates are the June 2028 figures from the original workbook until replaced by the project team.",
        "Fixtures and competing nations are not yet known; outputs are produced per risk level until they are.",
        "The project depends on the reference data in section 11 for several questions.",
        "Work is carried out on approved systems; files are not moved through personal accounts.",
    ])
    doc.add_heading("13. Risks", level=1)
    table(doc, ["Risk", "Impact", "Mitigation"], [
        ["Forces alter the form (insert rows, rename sheets) or use an old version.", "Figures mis-read or missing.",
         "Locked layout, guidance, Data Quality page flags old/altered forms."],
        ["Reference data (skillset mapping, schedule) not supplied in time.", "Questions S1-S7, R1-R7, B1 only "
         "partly answered.", "Track in section 11; dashboard shows totals without the mapping."],
        ["Business rules for conversions and scenarios not agreed.", "O2-O8 cannot be answered.", "Agree rules with "
         "portfolio leads; record decisions in the change log."],
        ["Inconsistent costing between mutual aid methods.", "Mis-stated budget.", "Decision 4; show both bases "
         "until agreed."],
        ["Single person dependency for the Power BI build.", "Delay if unavailable.", "Step-by-step build guide; "
         "documented queries and measures; test pack."],
    ], widths=[5.5, 4.5, 7])

    # ---------------------------------------------- 14 governance
    doc.add_heading("14. Governance, security and change control", level=1)
    bullets(doc, [
        ("Classification: ", "the form, returns, model and dashboard are OFFICIAL and handled in line with the HMG "
         "Security Policy Framework (risk management): https://www.gov.uk/government/publications/security-policy-"
         "framework/hmg-security-policy-framework#risk-management"),
        ("Freedom of Information: ", "requests are handled through npcc.foi.request@npfdu.police.uk."),
        ("Access: ", "Power BI workspace and returns folder restricted to named users in the client groups."),
        ("Change control: ", "changes to the form, calculation rules or reference data are logged with date, "
         "reason and approver; the form carries a version number."),
        ("Reporting: ", "progress reported to the project lead [frequency TBC]; issues and decisions recorded in "
         "section 11."),
    ])
    doc.add_heading("15. Milestones", level=1)
    table(doc, ["Milestone", "Target date"], [
        ["ToR agreed", "[TBC]"], ["Collection Form v3 signed off", "[TBC]"],
        ["Skillset mapping and fixture schedule supplied", "[TBC]"], ["Form issued to forces", "[TBC]"],
        ["Power BI model built and validated with the test pack", "[TBC]"], ["Returns received and loaded", "[TBC]"],
        ["Dashboard released to client groups", "[TBC]"], ["Update after the qualifying draw (teams, risk levels)",
                                                           "[TBC]"],
    ], widths=[11, 5])
    doc.add_heading("16. Approval", level=1)
    table(doc, ["Name", "Role", "Signature", "Date"], [["", "", "", ""], ["", "", "", ""]], widths=[4.5, 5, 4, 3])

    # ---------------------------------------------- appendices
    doc.add_heading("Appendix A - Glossary", level=1)
    table(doc, ["Term", "Meaning"], [
        ["MD / MD-2 / MD-1 / MD+1", "Match Day; two days before; the day before; the day after."],
        ["WT", "Wider Tournament - policing on tournament days outside the match-day cycle (e.g. fan zones, "
               "wider community)."],
        ["FS / MA", "Force Supplied / Mutual Aid Supplied."],
        ["Normal ToD", "Normal Tour of Duty hours per officer."],
        ["PSU / BDU", "Police Support Unit / Level 3 public order unit."],
        ["Risk level", "Low, Medium, High, Super High - the assumed risk of a fixture or tournament day."],
        ["Officer deployments", "Officers x days."],
        ["Master table", "The combined table Power BI builds from all returns: one row per force, resource, "
                         "origin, rank, risk level and day."],
        ["Reference table", "Data maintained by the planning team (rates, days, schedule, skillset mapping)."],
        ["ARV / CTSFO / AFO", "Armed Response Vehicle / Counter Terrorist Specialist Firearms Officer / Authorised "
                              "Firearms Officer."],
        ["CBRN / LFR / UAS", "Chemical, Biological, Radiological and Nuclear / Live Facial Recognition / Unmanned "
                             "Aerial Systems."],
    ], widths=[4.5, 12])
    doc.add_heading("Appendix B - Sections covered by the Collection Form (v3)", level=1)
    para(doc, "Match Day Resource", bold=True)
    para(doc, "Public Order (Level 2 PSU officers, Public Order Protective Carrier Drivers, Level 3 officers - by "
              "deployment location); Evidence Gathering; Protest Removal; Additional Public Order / Public Safety "
              "roles; Mounted; Dogs; Football Policing; Roads Policing; Armed Policing; Additional Armed Policing "
              "(Armed Response Overlay roles); Disruptive Effects; Search; Unmanned Aerial Systems (Drones); Live "
              "Facial Recognition; CBRN; Force Command; Control Room; Custody; Crime; Intelligence; Ancillary; "
              "Additional Resources.", size=9.5)
    para(doc, "Wider Tournament Resource", bold=True)
    para(doc, "Public Order (Wider Community / Fan Zone); Operationally Deployable non-Public Order Policing; "
              "Evidence Gatherers; Additional Public Order / Public Safety roles and Protest Removal; Fan Zone "
              "Disruptive Effects; Fan Zone Armed Policing; Wider Tournament Policing (Control Room, Command "
              "Support, Custody); Additional Resources.", size=9.5)
    doc.save(out)
    return out


if __name__ == "__main__":
    print(build(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
