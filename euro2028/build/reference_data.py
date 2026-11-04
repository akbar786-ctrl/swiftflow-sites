"""
Reference data and costing parameters.

Every value here is transcribed from the legacy workbook
"Base Assumptions MPSTH - Uniform.xlsx" (sheet + cell noted) or is a
decision recorded with the project team (decision ID noted). Nothing is
invented: where a value is unknown it is left as None and logged in
OPEN_ITEMS so it surfaces in the data-quality report.
"""
import datetime as dt

# ---------------------------------------------------------------------------
# Base hourly rates, financial year 2025-26
# ---------------------------------------------------------------------------
# rate_table: MA = "MUTUAL AID RATES - DO NOT EDIT"  (Baseline Costs!B1:H6)
#             HF = "HOST FORCE OVERTIME RATE"        (Baseline Costs!B18:H23)
# rate_type:  TOD = Hourly Employable Cost (Normal ToD hours)
#             OT  = Hourly Overtime Cost   (Extra Hours)
BASE_RATES = [
    # table, rank, TOD, OT, source TOD, source OT
    ("MA", "PC", 49.42, 49.42, "Baseline Costs!C5", "Baseline Costs!C6"),
    ("MA", "SGT", 61.49, 61.49, "Baseline Costs!D5", "Baseline Costs!D6"),
    ("MA", "INSP", 61.43, 61.43, "Baseline Costs!E5", "Baseline Costs!E6"),
    ("MA", "CI", 68.27, 68.27, "Baseline Costs!F5", "Baseline Costs!F6"),
    ("MA", "SUPT", 84.20, 84.20, "Baseline Costs!G5", "Baseline Costs!G6"),
    ("MA", "CSUPT", 97.21, 97.21, "Baseline Costs!H5", "Baseline Costs!H6"),
    ("HF", "PC", 49.42, 42.71, "Baseline Costs!C22", "Baseline Costs!C23"),
    ("HF", "SGT", 61.49, 51.29, "Baseline Costs!D22", "Baseline Costs!D23"),
    ("HF", "INSP", 61.43, 62.41, "Baseline Costs!E22", "Baseline Costs!E23"),
    ("HF", "CI", 68.27, 65.90, "Baseline Costs!F22", "Baseline Costs!F23"),
    ("HF", "SUPT", 84.20, 79.69, "Baseline Costs!G22", "Baseline Costs!G23"),
    ("HF", "CSUPT", 97.21, 92.74, "Baseline Costs!H22", "Baseline Costs!H23"),
]

# Annual employable costs (kept for reference / Command Team planning)
ANNUAL_COSTS = [
    ("MA", "PC", 50053), ("MA", "SGT", 63804), ("MA", "INSP", 76652),
    ("MA", "CI", 83959), ("MA", "SUPT", 103452), ("MA", "CSUPT", 119468),
    ("HF", "PC", 63788), ("HF", "SGT", 76597), ("HF", "INSP", 93202),
    ("HF", "CI", 98414), ("HF", "SUPT", 119023), ("HF", "CSUPT", 138502),
    ("CTP", "PC", 80802), ("CTP", "SGT", 97437), ("CTP", "INSP", 119275),
    ("CTP", "CI", 126079), ("CTP", "SUPT", 152971), ("CTP", "CSUPT", 178370),
]

# GDP deflators (GDP Deflators!B75:D83, ONS/OBR as at March 2026)
GDP_DEFLATORS = [
    ("2022-23", 91.2996, 7.029011671254197),
    ("2023-24", 96.1117, 5.27066931289951),
    ("2024-25", 100.0, 4.0456052697018166),
    ("2025-26", None, 3.238831156796862),
    ("2026-27", None, 2.0057419249593256),
    ("2027-28", None, 1.9419802624865978),
    ("2028-29", None, 1.8653157522025676),
    ("2029-30", None, 1.864367236103881),
    ("2030-31", None, 2.0159623211262856),
]
PAY_AWARDS = [("2025-26", 0.042), ("2026-27", 0.04), ("2027-28", 0.035), ("2028-29", 0.03)]

_defl_2627 = 1 + GDP_DEFLATORS[4][2] / 100   # Baseline Costs!K4 = 1.0200574...

# Uplift steps applied cumulatively to the 2025-26 base, replicating the
# legacy method exactly (Decision D-04: rate tables are final).
#   MA rates  Sep-26: rate*K4 + rate*N4  -> factor (K4 + N4)   [additive]
#             Sep-27: *(1+N5)            Sep-28: *(1+N6)
#   HF rates  Sep-26: rate + rate*N21    -> factor (1 + N21)    [no deflator]
#   MA PSU    Sep-26: Q6 + Q6*N4         -> factor (1 + N4)
UPLIFT_STEPS = [
    # table, effective_from, factor, method / source
    ("MA", dt.date(2026, 9, 1), _defl_2627 + 0.04, "GDP deflator index 26/27 + pay award 26/27 (Baseline Costs!C10:C15)"),
    ("MA", dt.date(2027, 9, 1), 1.035, "Pay award 27/28 (Baseline Costs!O10:O15)"),
    ("MA", dt.date(2028, 9, 1), 1.03, "Pay award 28/29 (Baseline Costs!AA10:AA15)"),
    ("HF", dt.date(2026, 9, 1), 1.04, "Pay award 26/27 (Baseline Costs!C27:C32)"),
    ("HF", dt.date(2027, 9, 1), 1.035, "Pay award 27/28 (Baseline Costs!O27:O32)"),
    ("HF", dt.date(2028, 9, 1), 1.03, "Pay award 28/29 (Baseline Costs!AA27:AA32)"),
    ("CTP", dt.date(2026, 9, 1), _defl_2627 + 0.04, "Monthly salary x (deflator 26/27 + pay award 26/27) (Command Team & Planning!B13)"),
    ("CTP", dt.date(2027, 9, 1), 1.035, "Pay award 27/28 (Command Team & Planning!N13)"),
    ("CTP", dt.date(2028, 9, 1), 1.03, "Pay award 28/29 (Command Team & Planning!Z13)"),
    ("MAPSU", dt.date(2026, 9, 1), 1.04, "Pay award 26/27 (Baseline Costs!C17)"),
    ("MAPSU", dt.date(2027, 9, 1), 1.035, "Pay award 27/28 (Baseline Costs!O17)"),
    ("MAPSU", dt.date(2028, 9, 1), 1.03, "Pay award 28/29 (Baseline Costs!AA17)"),
]

# Mutual-aid PSU package (Baseline Costs!P3:Q6). Decision D-06: covers
# travel day in, the match day worked, and travel day out.
MA_PSU_PACKAGE = [
    # unit type, base 2025-26 value, note
    ("Level 2 PSU", 31601.07 * 1.05, "IDR/TDR SUB 31,601.07 + 5% admin (Baseline Costs!Q4:Q6)"),
    ("Level 3 BDU", 31601.07 * 1.05, "Legacy applied the same package to Level 3 BDU - OPEN ITEM OI-03"),
    ("Level 2 Serial", None, "No rate in legacy workbook - OPEN ITEM OI-03"),
    ("Level 3 Serial", None, "No rate in legacy workbook - OPEN ITEM OI-03"),
]

# ---------------------------------------------------------------------------
# Rank costing rules
# ---------------------------------------------------------------------------
# Decision D-05: force-supplied Inspectors are GBP 0 for ToD and Extra Hours.
# Ranks above Inspector, force supplied: GBP 0 in legacy (no rate) - assumed
# the same, logged as OI-02. Mutual-aid Inspector+ are charged at the MA rate
# as in legacy (OI-02).
ZERO_COST_RULES = [
    # supply, rank, applies_to, decision
    ("FS", "INSP", "TOD+OT", "D-05"),
    ("FS", "CI", "TOD+OT", "Legacy (no rate) - OI-02"),
    ("FS", "SUPT", "TOD+OT", "Legacy (no rate) - OI-02"),
    ("FS", "CSUPT", "TOD+OT", "Legacy (no rate) - OI-02"),
    ("FS", "ACC", "TOD+OT", "Legacy (no rate) - OI-02"),
    ("MA", "ACC", "TOD+OT", "No ACC rate in either table - OI-02"),
]

# ---------------------------------------------------------------------------
# Parameters (scenario / national assumptions)
# ---------------------------------------------------------------------------
PARAMETERS = [
    # name, value, type, description, source
    ("RuleSet", "LEGACY", "text",
     "LEGACY reproduces the old workbook's calculation logic (selected by the project team). AGREED applies "
     "the revised rules D-02, D-03, D-05, D-06, D-07. See docs/02_calculation_rules.md.", "Decision D-12"),
    ("OtherResourcesRateDate2728", dt.date(2027, 6, 1), "date",
     "Rate month for Other Resources FY 2027-28 hours.", "Legacy Other Resources!E7 used Baseline Costs column L (Jun-2027)."),
    ("RateEffectiveDate", dt.date(2028, 6, 1), "date",
     "Date at which hourly rates and the MA package are priced (tournament month).",
     "Legacy Cost Calculator used Baseline Costs column X (Jun-2028)."),
    ("OverlapRule", "ADD", "text",
     "When one date is several relative days for different fixtures, demands are ADDED.", "Decision D-02"),
    ("MD1_PSU_Factor", 1.0, "number",
     "Factor applied to PSU counts on MD+1. Legacy halved PSUs on some dates only; not adopted.",
     "Decision D-02 (separate categories, add both)"),
    ("MAPSU_Basis_Matchday", "PER_UNIT_PER_FIXTURE", "text",
     "Matchday-only mutual-aid PSUs: one package per unit per fixture.", "Decision D-06"),
    ("MAPSU_Basis_NonMatchday", "PER_UNIT_PER_FIXTURE_MAXDAY", "text",
     "Non-matchday mutual-aid PSUs: one package per unit, using the peak day count across MD-2..MD+1. "
     "Alternative PER_UNIT_PER_DAY reproduces legacy.", "OPEN ITEM OI-03"),
    ("MAPSU_Basis_WT", "PER_DAY", "text",
     "Wider tournament mutual-aid PSUs: one package per unit per typical day, multiplied by the national "
     "WT day assumptions (reproduces legacy WT Cost Calculator rows 303-311).", "OPEN ITEM OI-03"),
    ("MA_ToD_Collected", "FALSE", "bool",
     "Whether Normal ToD is collected/costed on mutual-aid officer rows.", "OPEN ITEM OI-01"),
    ("TemplateVersion", "2.0.0", "text", "Collection template version accepted by the model.", "Build"),
]

# Legacy "Other Resources" categories (Other Resources!A6:A26), each with Sergeant and Constable rows
OTHER_RESOURCE_CATEGORIES = ["Media Management", "Licensing", "Crime Investigation", "Intelligence Arrangements",
                             "Post Event Investigation", "Professional Standards & Legal", "Command Team & Planning"]

WT_DAY_ASSUMPTIONS = [  # risk code, days, source  (Decision D-08: national assumptions)
    ("L", 7, "WT Cost Template!H7"), ("M", 11, "WT Cost Template!I7"),
    ("H", 11, "WT Cost Template!J7"), ("SH", 3, "WT Cost Template!K7"),
]

# ---------------------------------------------------------------------------
# Forces (region assignment to be confirmed by project team - OI-07)
# ---------------------------------------------------------------------------
FORCES = [
    ("Avon and Somerset Constabulary", "South West"), ("Bedfordshire Police", "Eastern"),
    ("Cambridgeshire Constabulary", "Eastern"), ("Cheshire Constabulary", "North West"),
    ("City of London Police", "London"), ("Cleveland Police", "North East"),
    ("Cumbria Constabulary", "North West"), ("Derbyshire Constabulary", "East Midlands"),
    ("Devon and Cornwall Police", "South West"), ("Dorset Police", "South West"),
    ("Durham Constabulary", "North East"), ("Dyfed-Powys Police", "Wales"),
    ("Essex Police", "Eastern"), ("Gloucestershire Constabulary", "South West"),
    ("Greater Manchester Police", "North West"), ("Gwent Police", "Wales"),
    ("Hampshire and Isle of Wight Constabulary", "South East"), ("Hertfordshire Constabulary", "Eastern"),
    ("Humberside Police", "Yorkshire and the Humber"), ("Kent Police", "South East"),
    ("Lancashire Constabulary", "North West"), ("Leicestershire Police", "East Midlands"),
    ("Lincolnshire Police", "East Midlands"), ("Merseyside Police", "North West"),
    ("Metropolitan Police Service", "London"), ("Norfolk Constabulary", "Eastern"),
    ("North Wales Police", "Wales"), ("North Yorkshire Police", "Yorkshire and the Humber"),
    ("Northamptonshire Police", "East Midlands"), ("Northumbria Police", "North East"),
    ("Nottinghamshire Police", "East Midlands"), ("South Wales Police", "Wales"),
    ("South Yorkshire Police", "Yorkshire and the Humber"), ("Staffordshire Police", "West Midlands"),
    ("Suffolk Constabulary", "Eastern"), ("Surrey Police", "South East"),
    ("Sussex Police", "South East"), ("Thames Valley Police", "South East"),
    ("Warwickshire Police", "West Midlands"), ("West Mercia Police", "West Midlands"),
    ("West Midlands Police", "West Midlands"), ("West Yorkshire Police", "Yorkshire and the Humber"),
    ("Wiltshire Police", "South West"),
    ("Police Scotland", "Scotland"), ("Police Service of Northern Ireland", "Northern Ireland"),
    ("British Transport Police", "National"), ("Civil Nuclear Constabulary", "National"),
    ("Ministry of Defence Police", "National"),
]

# ---------------------------------------------------------------------------
# Decisions and open items
# ---------------------------------------------------------------------------
DECISIONS = [
    ("D-01", "Risk level per fixture is an assumption; forces complete every risk level. Fixtures are not "
             "known, so the model works per fixture and expands to a calendar only when a fixture list exists."),
    ("D-02", "Where one date is more than one relative day (e.g. MD+1 of one fixture and MD-2 of the next), "
             "the demands are ADDED. MD and MD+1 are separate categories."),
    ("D-03", "Normal ToD hours are costed for all officers (legacy costed Extra Hours only)."),
    ("D-04", "Host Force and Mutual Aid rate tables are final and used exactly as provided."),
    ("D-05", "Inspectors are GBP 0, for Normal ToD and Extra Hours."),
    ("D-06", "The GBP 35,716 mutual-aid PSU package covers travel day in, match day and travel day out."),
    ("D-07", "Additional resources and role-profile rows are priced by the rank selected."),
    ("D-08", "Wider Tournament keeps Super High. WT day counts per risk level are national assumptions."),
    ("D-09", "Level 1 PSU, Servator Complex/Non-Complex, CTSFO Overlay and MA Command are retired."),
    ("D-10", "Command Team & Planning, Other Resources and Fleet are brought into Power BI."),
    ("D-11", "Template layout: forces' form (columns A-C), risk level on top, relative day beneath, measure "
             "beneath that. Super High is added to Match Day."),
    ("D-12", "Calculations reproduce the legacy workbook logic (RuleSet = LEGACY), in Power BI and in the backup "
             "calculation workbook. The agreed revised rules remain available as RuleSet = AGREED."),
]

OPEN_ITEMS = [
    ("OI-01", "Mutual-aid Normal ToD",
     "The draft template hatches Normal ToD on mutual-aid rows, but D-03 says ToD is costed for all officers. "
     "Model default follows the template (not collected). Switch: parameter MA_ToD_Collected.", "Open"),
    ("OI-02", "Costing of ranks above Inspector",
     "Force-supplied CI/Supt/C Supt/ACC assumed GBP 0 (as legacy). Mutual-aid Inspector+ charged at MA rate "
     "(as legacy). Confirm both.", "Open"),
    ("OI-03", "Mutual-aid PSU package basis",
     "'Mutual aid costs are different' - confirm (a) number of packages for non-matchday MA PSUs spanning "
     "several days, (b) Level 3 BDU price, (c) serial price.", "Open"),
    ("OI-04", "Mercury skillset mapping",
     "Draft mapping of each resource to a Mercury skill is provided for confirmation.", "Open"),
    ("OI-05", "Equipment ratios",
     "Protected carriers per L2 PSU, horses per mounted officer, C-UAS kit, LFR vans - ratios required. "
     "Legacy Fleet sheet multiplied by 3 (meaning unknown).", "Open"),
    ("OI-06", "Scenario uplifts",
     "Op Brock, Op Stack, Israel qualifying and national policing reserve assumptions required.", "Open"),
    ("OI-07", "Force regions and capacity",
     "Region per force (draft provided) and available capacity per force/skill/rank required.", "Open"),
    ("OI-08", "Role profiles vs dedicated rows",
     "Protest Removal exists as dedicated rows and as role profiles under Additional Public Order - guidance "
     "must say which to use to avoid double counting.", "Open"),
]
