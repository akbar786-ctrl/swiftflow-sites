"""
EURO 2028 resource catalogue - single source of truth.

Every collection row, dropdown, N/A rule, reference table and Power BI
dimension is generated from the definitions in this file. Change a resource
here, re-run build_all.py, and the template, reference workbook and mapping
are regenerated consistently.

Row types
---------
PSU_CONFIG   Officers per ONE PSU/serial, by rank, with hours per officer.
PSU_COUNT    Number of force-supplied PSUs/serials (units). Hours come from
             the linked PSU_CONFIG resource.
MA_PSU       Number of mutual-aid PSUs/serials (units). Costed by package.
OFFICER      Fixed rank per line.
RANK_SELECT  Rank chosen from a dropdown on each line.
ROLE         Role profile + rank chosen from dropdowns on each line.
ADDITIONAL   Free-text resource + supply type + rank chosen by the force.
"""

from collections import OrderedDict

TEMPLATE_VERSION = "2.0.0"

# Sheet names match the forces' draft form
SHEET_NAMES = {"MD": "Match Day Resource", "WT": "Wider Tournament Resource"}
LISTS_SHEET = "Data Info"

# ---------------------------------------------------------------------------
# Dimensions
# ---------------------------------------------------------------------------
RISK_LEVELS = [  # code, label, sort
    ("L", "Low", 1),
    ("M", "Medium", 2),
    ("H", "High", 3),
    ("SH", "Super High", 4),
]

MD_DAYS = [  # code, label, offset from match day, sort
    ("MD-2", "Match Day -2", -2, 1),
    ("MD-1", "Match Day -1", -1, 2),
    ("MD", "Match Day", 0, 3),
    ("MD+1", "Match Day +1", 1, 4),
]

MEASURES = [  # code, label
    ("NUM", "Number required"),
    ("TOD", "Normal ToD (hrs)"),
    ("XH", "Extra Hours (hrs)"),
]

# Rank master. rate_rank = the rank used to look up a rate.
RANKS = [  # code, label, rate_rank, sort, is_detective
    ("ACC", "ACC", "ACC", 1, False),
    ("CSUPT", "Chief Superintendent", "CSUPT", 2, False),
    ("SUPT", "Superintendent", "SUPT", 3, False),
    ("CI", "Chief Inspector", "CI", 4, False),
    ("INSP", "Inspector", "INSP", 5, False),
    ("SGT", "Sergeant", "SGT", 6, False),
    ("PC", "Constable", "PC", 7, False),
    ("DI", "Detective Inspector", "INSP", 5, True),
    ("DS", "Detective Sergeant", "SGT", 6, True),
    ("DC", "Detective Constable", "PC", 7, True),
]
RANK_LABEL = {r[0]: r[1] for r in RANKS}
RANK_BY_LABEL = {r[1]: r[0] for r in RANKS}

# Dropdown rank lists (list name -> rank codes)
RANK_LISTS = OrderedDict([
    ("Ranks_Std", ["INSP", "SGT", "PC"]),
    ("Ranks_SgtPC", ["SGT", "PC"]),
    ("Ranks_All", ["ACC", "CSUPT", "SUPT", "CI", "INSP", "SGT", "PC"]),
    ("Ranks_Det", ["DI", "DS", "DC"]),
])

UNIT_TYPES = [  # code, label, psu_level, is_serial
    ("L2PSU", "Level 2 PSU", "L2", False),
    ("L2SER", "Level 2 Serial", "L2", True),
    ("L3BDU", "Level 3 BDU", "L3", False),
    ("L3SER", "Level 3 Serial", "L3", True),
]
UNIT_LISTS = OrderedDict([
    ("Units_L2", ["Level 2 PSU", "Level 2 Serial"]),
    ("Units_L3", ["Level 3 BDU", "Level 3 Serial"]),
])

# From the new template's Data Info sheet (typo "Renhanced" corrected).
ROLE_PROFILES = [  # portfolio, role profile
    ("Armed Policing", "Firearms Tactical Advisor"),
    ("Armed Policing", "Specialist Firearms Officer"),
    ("Armed Policing", "Specialist Strategic Firearms Officer"),
    ("Armed Policing", "Specialist Tactical Firearms Officer"),
    ("Armed Policing", "Strategic Firearms Commander"),
    ("Armed Policing", "ARV"),
    ("Armed Policing", "CTSFO"),
    ("CBRN", "CBRN - Enhanced Responder"),
    ("CBRN", "CBRN - Operational Commander (Bronze)"),
    ("CBRN", "CBRN - Search"),
    ("CBRN", "CBRN - Tactical Advisor"),
    ("CBRN", "CBRN - Tactical Commander (Silver)"),
    ("Public Order / Public Safety", "Forward Intelligence Team Officer (FIT)"),
    ("Public Order / Public Safety", "Police Liaison Officer (PLT)"),
    ("Public Order / Public Safety", "Medic"),
    ("Public Order / Public Safety", "PO Bronze Commander"),
    ("Public Order / Public Safety", "PO Safety Advisor"),
    ("Public Order / Public Safety", "PO Loggist"),
    ("Public Order / Public Safety", "Non-Complex Protest Removal"),
    ("Public Order / Public Safety", "Working at Height - Protest Removal"),
    ("Public Order / Public Safety", "De-Bonding - Protest Removal"),
]
ROLE_LISTS = OrderedDict([
    ("Roles_Armed", [r for p, r in ROLE_PROFILES if p == "Armed Policing"]),
    ("Roles_CBRN", [r for p, r in ROLE_PROFILES if p == "CBRN"]),
    ("Roles_POPS", [r for p, r in ROLE_PROFILES if p == "Public Order / Public Safety"]),
])

WT_OPTIONS = ["Wider Community", "Fan Zone"]
SUPPLY_TYPES = [("FS", "Force Supplied"), ("MA", "Mutual Aid Requested")]
SECTION_STATUS = ["Not started", "In progress", "Complete", "Not required"]

# ---------------------------------------------------------------------------
# Resource definitions
# ---------------------------------------------------------------------------
# Compact helper. Each resource -> dict. Lines are generated from ranks/n_lines.
_RES = []


def _r(coll, sec, num, name, supply, rtype, portfolio, location, *,
       ranks=None, n_lines=None, rank_list=None, role_list=None,
       units=None, psu_config=None, matchday_only=False, basis="Officers",
       mercury=None, instructions=None):
    rid = f"{coll}-{sec}-{num:02d}"
    # Legacy Demand Sheet priority group (which relative day wins when a date is
    # several relative days for fixtures of the same risk level):
    #   MDONLY  matchday-only resources: Match Day only
    #   NONMD   non-matchday public order: MD-1 > MD-2 > MD+1 (Match Day not used)
    #   STD     all other resources: MD > MD-1 > MD-2 > MD+1
    day_group = "MDONLY" if matchday_only else ("NONMD" if location == "Non-Matchday" else "STD")
    _RES.append(dict(
        resource_id=rid, collection=coll, section_code=sec, name=name,
        supply=supply, row_type=rtype, portfolio=portfolio, location=location,
        ranks=ranks, n_lines=n_lines, rank_list=rank_list, role_list=role_list,
        units=units, psu_config=psu_config, matchday_only=matchday_only,
        basis=basis, mercury=mercury or "", instructions=instructions or "",
        day_group=day_group,
        # Legacy Cost Calculator charged In-Stadium PSU Inspectors at the MA Inspector rate
        legacy_insp_ma=(rtype == "PSU_CONFIG" and location == "In Stadium"),
    ))
    return rid


STD = ["INSP", "SGT", "PC"]
SP = ["SGT", "PC"]
CFG_TXT = "Enter the number of each rank in ONE {lvl} / serial, with the hours each officer works."
CNT_TXT = "Enter the number of PSUs / serials. Choose PSU or Serial in column C."

# Sections: code -> (collection, title, portfolio-for-header)
SECTIONS = OrderedDict()


def _section(coll, code, title):
    SECTIONS[(coll, code)] = title


# ---------------- MATCH DAY ----------------
_section("MD", "PSC", "PSU / Serial Configurations")
cfg = {}
for i, (loc, lbl, mdo) in enumerate([
        ("Non-Matchday", "(Non-Matchday)", False),
        ("Town Centre", "Town Centre Public Order Policing", True),
        ("Stadium Vicinity", "Stadium Vicinity Public Order Policing", True),
        ("In Stadium", "In Stadium Public Order Policing", True)]):
    for j, lvl in enumerate(["L2", "L3"]):
        unit = "Level 2 PSU" if lvl == "L2" else "Level 3 BDU"
        nm = f"Force Supplied - {'Level 2 PSU' if lvl == 'L2' else 'Level 3 BDU'} / Serials {lbl}"
        cfg[(loc, lvl)] = _r("MD", "PSC", i * 2 + j + 1, nm, "FS", "PSU_CONFIG",
                             "Public Order / Public Safety", loc, ranks=STD,
                             matchday_only=mdo, basis="Officers per unit",
                             mercury=f"PO {'Level 2' if lvl == 'L2' else 'Level 3'}",
                             instructions=CFG_TXT.format(lvl=unit))

_section("MD", "POP", "Public Order / Public Safety")
n = 0
for loc, lbl, mdo in [("Non-Matchday", "Public Order (Non-Matchday)", False),
                      ("Town Centre", "Town Centre Public Order Policing (Matchday Only)", True),
                      ("Stadium Vicinity", "Stadium Vicinity Public Order Policing (Matchday Only)", True),
                      ("In Stadium", "In Stadium Public Order Policing (Matchday Only)", True)]:
    n += 1
    _r("MD", "POP", n, f"Force Supplied - {lbl}", "FS", "PSU_COUNT",
       "Public Order / Public Safety", loc, units=["L2", "L3"],
       psu_config={"L2": cfg[(loc, "L2")], "L3": cfg[(loc, "L3")]},
       matchday_only=mdo, basis="Units", mercury="PO Level 2 / Level 3", instructions=CNT_TXT)
    n += 1
    _r("MD", "POP", n, f"Mutual Aid Requested - {lbl}", "MA", "MA_PSU",
       "Public Order / Public Safety", loc, units=["L2", "L3"],
       matchday_only=mdo, basis="Units", mercury="PO Level 2 / Level 3", instructions=CNT_TXT)

_section("MD", "EGT", "Evidence Gathering Team Officer")
_r("MD", "EGT", 1, "Force Supplied - Evidence Gathering Team Officer", "FS", "OFFICER",
   "Public Order / Public Safety", "General", ranks=SP, mercury="EGT Officer")
_r("MD", "EGT", 2, "Mutual Aid Requested - Evidence Gathering Team Officer", "MA", "OFFICER",
   "Public Order / Public Safety", "General", ranks=SP, mercury="EGT Officer")

_section("MD", "PRT", "Protest Removal")
for k, (nm, merc) in enumerate([("Working at Height Complex Protest Removal", "PRT Complex"),
                                ("Non-Complex Protest Removal", "PRT Non Complex"),
                                ("Debonding Protest Removal", "PRT Debonding")]):
    _r("MD", "PRT", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Public Order / Public Safety", "General", ranks=STD, mercury=merc)
    _r("MD", "PRT", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Public Order / Public Safety", "General", ranks=STD, mercury=merc)

_section("MD", "APO", "Additional Public Order / Public Safety Resources")
_r("MD", "APO", 1, "Force Supplied - Additional Public Order", "FS", "ROLE",
   "Public Order / Public Safety", "General", n_lines=7, rank_list="Ranks_All",
   role_list="Roles_POPS", instructions="Choose the rank in column A and the role profile in column C.")
_r("MD", "APO", 2, "Mutual Aid Requested - Additional Public Order", "MA", "ROLE",
   "Public Order / Public Safety", "General", n_lines=6, rank_list="Ranks_All",
   role_list="Roles_POPS", instructions="Choose the rank in column A and the role profile in column C.")

_section("MD", "MNT", "Mounted")
for k, (nm, merc) in enumerate([("Mounted Commander", "Mounted - Commander"),
                                ("Mounted Officer", "Mounted - PO Officer")]):
    _r("MD", "MNT", k * 2 + 1, f"Force Supplied - {nm}", "FS", "RANK_SELECT",
       "Mounted", "General", n_lines=1, rank_list="Ranks_Std", mercury=merc)
    _r("MD", "MNT", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "RANK_SELECT",
       "Mounted", "General", n_lines=1, rank_list="Ranks_Std", mercury=merc)

_section("MD", "DOG", "Dogs")
for k, (nm, merc) in enumerate([("Dogs GDP", "Dogs - GPD"),
                                ("Dogs Public Order Serial Commander", "Dogs - PO Dog Serial Commander"),
                                ("Dogs Public Order", "Dogs - PO Dog"),
                                ("Dogs Explo", "Dogs - Explo Dog")]):
    _r("MD", "DOG", k * 2 + 1, f"Force Supplied - {nm}", "FS", "RANK_SELECT",
       "Dogs", "General", n_lines=3, rank_list="Ranks_Std", mercury=merc)
    _r("MD", "DOG", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "RANK_SELECT",
       "Dogs", "General", n_lines=3, rank_list="Ranks_Std", mercury=merc)

_section("MD", "FBP", "Football Policing Portfolio")
for k, (nm, merc) in enumerate([("Football Spotting DFO", "Football - DFO"),
                                ("Football Spotting OFO", "Football - OFO")]):
    _r("MD", "FBP", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Football Policing", "General", ranks=SP, mercury=merc)
    _r("MD", "FBP", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Football Policing", "General", ranks=SP, mercury=merc)

_section("MD", "RDS", "Roads Policing")
for k, (nm, merc) in enumerate([("Roads Policing Officer", "Roads Policing"),
                                ("Roads Policing Officer - Motorcyclists", "Roads - Motorcyclist"),
                                ("Roads Policing Officer - VIP Motorcyclist Escort", "Roads - VIPEX")]):
    _r("MD", "RDS", k * 2 + 1, f"Force Supplied - {nm}", "FS", "RANK_SELECT",
       "Roads Policing", "General", n_lines=2, rank_list="Ranks_Std", mercury=merc)
    _r("MD", "RDS", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "RANK_SELECT",
       "Roads Policing", "General", n_lines=2, rank_list="Ranks_Std", mercury=merc)

_section("MD", "ARM", "Armed Policing")
for k, (nm, rk, merc) in enumerate([("Operational Firearms Commander", STD, "Firearms - OFC"),
                                    ("Armed Response (ARVs)", STD, "Firearms - ARV"),
                                    ("CTSFOs", STD, "Firearms - CTSFO"),
                                    ("AFOs", STD, "Firearms - AFO"),
                                    ("Rifle Officer", SP, "Firearms - Rifle Officer"),
                                    ("Protection Officer", SP, "Firearms - Protection Officer")]):
    _r("MD", "ARM", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Armed Policing", "General", ranks=rk, mercury=merc)
    _r("MD", "ARM", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Armed Policing", "General", ranks=rk, mercury=merc)

_section("MD", "AAP", "Additional Armed Policing")
_r("MD", "AAP", 1, "Force Supplied - Armed Response Overlay", "FS", "ROLE",
   "Armed Policing", "General", n_lines=7, rank_list="Ranks_All", role_list="Roles_Armed",
   instructions="Choose the rank in column A and the role profile in column C.")
_r("MD", "AAP", 2, "Mutual Aid Requested - Armed Response Overlay", "MA", "ROLE",
   "Armed Policing", "General", n_lines=7, rank_list="Ranks_All", role_list="Roles_Armed",
   instructions="Choose the rank in column A and the role profile in column C.")

_section("MD", "DFX", "Disruptive Effects")
for k, (nm, merc) in enumerate([("Servator Disruptive Effects", "DEO Officer"),
                                ("Servator Disruptive Effects Operational Commander", "DEO Commander"),
                                ("Servator Disruptive Effects Tactical Commander", "DEO Commander")]):
    _r("MD", "DFX", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Disruptive Effects", "General", ranks=SP, mercury=merc)
    _r("MD", "DFX", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Disruptive Effects", "General", ranks=SP, mercury=merc)

_section("MD", "SRC", "Search")
for k, (nm, merc) in enumerate([("LSOs", "Search - LSO"), ("PolSAs", "Search - PolSA"),
                                ("CT Secco", "CT SecCO")]):
    _r("MD", "SRC", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Search", "General", ranks=STD, mercury=merc)
    _r("MD", "SRC", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Search", "General", ranks=STD, mercury=merc)

_section("MD", "UAS", "Unmanned Aerial Systems (Drones)")
for k, (nm, merc) in enumerate([("Counter UAS Bronze Commander", "C-UAS - Bronze"),
                                ("Counter UAS Effector Operator", "C-UAS - Officer"),
                                ("Counter UAS Observers", "C-UAS - Officer"),
                                ("Counter UAS Tactical Advisor", "C-UAS - Tactical Advisor"),
                                ("Drone", "Drone Pilot")]):
    _r("MD", "UAS", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Drones (UAS)", "General", ranks=STD, mercury=merc)
    _r("MD", "UAS", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Drones (UAS)", "General", ranks=STD, mercury=merc)

_section("MD", "LFR", "Live Facial Recognition")
for k, (nm, merc) in enumerate([("LFR Operational Commander", "LFR - Commander"),
                                ("LFR Operator", "LFR - Operator")]):
    _r("MD", "LFR", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER",
       "Live Facial Recognition", "General", ranks=STD, mercury=merc)
    _r("MD", "LFR", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER",
       "Live Facial Recognition", "General", ranks=STD, mercury=merc)

_section("MD", "CBR", "CBRN Portfolio")
_r("MD", "CBR", 1, "Force Supplied - General Responder", "FS", "OFFICER", "CBRN", "General",
   ranks=STD, mercury="CBRN - General Responder")
_r("MD", "CBR", 2, "Mutual Aid Requested - General Responder", "MA", "OFFICER", "CBRN", "General",
   ranks=STD, mercury="CBRN - General Responder")
_r("MD", "CBR", 3, "Force Supplied - CBRN Officer(s)", "FS", "ROLE", "CBRN", "General",
   n_lines=5, rank_list="Ranks_All", role_list="Roles_CBRN",
   instructions="Choose the rank in column A and the CBRN role profile in column C.")
_r("MD", "CBR", 4, "Mutual Aid Requested - CBRN Officer(s)", "MA", "ROLE", "CBRN", "General",
   n_lines=5, rank_list="Ranks_All", role_list="Roles_CBRN",
   instructions="Choose the rank in column A and the CBRN role profile in column C.")

_section("MD", "CMD", "Non Public Order Force Command")
_r("MD", "CMD", 1, "Force Supplied - Command", "FS", "OFFICER", "Command", "General",
   ranks=["ACC", "CSUPT", "SUPT", "CI", "INSP", "SGT", "PC"], mercury="Command")

_section("MD", "CTL", "Control Room")
_r("MD", "CTL", 1, "Force Supplied - Control Room", "FS", "OFFICER", "Control Room", "General",
   ranks=STD, mercury="Control Room")

_section("MD", "CUS", "Custody")
_r("MD", "CUS", 1, "Force Supplied - Custody", "FS", "OFFICER", "Custody", "General",
   ranks=SP, mercury="Custody")

_section("MD", "CRM", "Crime")
_r("MD", "CRM", 1, "Force Supplied - Crime", "FS", "OFFICER", "Crime", "General",
   ranks=["DI", "DS", "DC"], mercury="Crime")
_r("MD", "CRM", 2, "Mutual Aid Requested - Crime", "MA", "OFFICER", "Crime", "General",
   ranks=["DI", "DS", "DC"], mercury="Crime")

_section("MD", "INT", "Intel")
_r("MD", "INT", 1, "Force Supplied - Intel", "FS", "OFFICER", "Intel", "General",
   ranks=STD, mercury="Intel")
_r("MD", "INT", 2, "Mutual Aid Requested - Intel", "MA", "OFFICER", "Intel", "General",
   ranks=STD, mercury="Intel")

_section("MD", "ANC", "Ancillary")
_r("MD", "ANC", 1, "Force Supplied - Ancillary Team", "FS", "OFFICER", "Ancillary", "General",
   ranks=SP, mercury="Ancillary")
_r("MD", "ANC", 2, "Mutual Aid Requested - Ancillary Team", "MA", "OFFICER", "Ancillary", "General",
   ranks=SP, mercury="Ancillary")

_section("MD", "ADD", "Additional Resources")
_r("MD", "ADD", 1, "Force Supplied - Additional Resource", "FS", "ADDITIONAL",
   "Additional", "General", n_lines=5, rank_list="Ranks_All",
   instructions="Describe the resource in column B and choose the rank in column C.")
_r("MD", "ADD", 2, "Mutual Aid Requested - Additional Resource", "MA", "ADDITIONAL",
   "Additional", "General", n_lines=5, rank_list="Ranks_All",
   instructions="Describe the resource in column B and choose the rank in column C.")

# ---------------- WIDER TOURNAMENT ----------------
_section("WT", "PSC", "PSU / Serial Configurations")
wcfg = {}
for i, (loc, lbl) in enumerate([("Wider Community", "Wider Community Public Order Policing"),
                                ("Fan Zone", "Fan Zone Public Order Policing")]):
    for j, lvl in enumerate(["L2", "L3"]):
        unit = "Level 2 PSU" if lvl == "L2" else "Level 3 BDU"
        wcfg[(loc, lvl)] = _r("WT", "PSC", i * 2 + j + 1, f"Force Supplied - {unit} / Serials {lbl}",
                              "FS", "PSU_CONFIG", "Public Order / Public Safety", loc, ranks=STD,
                              basis="Officers per unit",
                              mercury=f"PO {'Level 2' if lvl == 'L2' else 'Level 3'}",
                              instructions=CFG_TXT.format(lvl=unit))

_section("WT", "POP", "Public Order / Public Safety")
_r("WT", "POP", 1, "Force Supplied - Wider Community Public Order Policing", "FS", "PSU_COUNT",
   "Public Order / Public Safety", "Wider Community", units=["L2", "L3"],
   psu_config={"L2": wcfg[("Wider Community", "L2")], "L3": wcfg[("Wider Community", "L3")]},
   basis="Units", mercury="PO Level 2 / Level 3", instructions=CNT_TXT)
_r("WT", "POP", 2, "Mutual Aid Requested - Wider Community Public Order Policing", "MA", "MA_PSU",
   "Public Order / Public Safety", "Wider Community", units=["L2", "L3"], basis="Units",
   mercury="PO Level 2 / Level 3", instructions=CNT_TXT)
_r("WT", "POP", 3, "Force Supplied - Fan Zone Public Order Policing", "FS", "PSU_COUNT",
   "Public Order / Public Safety", "Fan Zone", units=["L2", "L3"],
   psu_config={"L2": wcfg[("Fan Zone", "L2")], "L3": wcfg[("Fan Zone", "L3")]},
   basis="Units", mercury="PO Level 2 / Level 3", instructions=CNT_TXT)
_r("WT", "POP", 4, "Mutual Aid Requested - Fan Zone Public Order Policing", "MA", "MA_PSU",
   "Public Order / Public Safety", "Fan Zone", units=["L2", "L3"], basis="Units",
   mercury="PO Level 2 / Level 3", instructions=CNT_TXT)
_r("WT", "POP", 5, "Force Supplied - Operationally Deployable non-Public Order Policing", "FS",
   "OFFICER", "Public Order / Public Safety", "Wider Community", ranks=STD, mercury="GPD")
_r("WT", "POP", 6, "Mutual Aid Requested - Operationally Deployable non-Public Order Policing", "MA",
   "OFFICER", "Public Order / Public Safety", "Wider Community", ranks=STD, mercury="GPD")
_r("WT", "POP", 7, "Force Supplied - Evidence Gathering Team Officer", "FS", "OFFICER",
   "Public Order / Public Safety", "General", ranks=SP, basis="Shifts", mercury="EGT Officer",
   instructions="Number of shifts")
_r("WT", "POP", 8, "Mutual Aid Requested - Evidence Gathering Team Officer", "MA", "OFFICER",
   "Public Order / Public Safety", "General", ranks=SP, basis="Shifts", mercury="EGT Officer",
   instructions="Number of shifts")

_section("WT", "APO", "Additional POPS Resources")
for k, (loc, nf, nm) in enumerate([("Wider Community", 6, 5), ("Fan Zone", 5, 4)]):
    _r("WT", "APO", k * 2 + 1, f"Force Supplied - Additional Public Order ({loc})", "FS", "ROLE",
       "Public Order / Public Safety", loc, n_lines=nf, rank_list="Ranks_All", role_list="Roles_POPS",
       instructions="Choose the rank in column A and the role profile in column C.")
    _r("WT", "APO", k * 2 + 2, f"Mutual Aid Requested - Additional Public Order ({loc})", "MA", "ROLE",
       "Public Order / Public Safety", loc, n_lines=nm, rank_list="Ranks_All", role_list="Roles_POPS",
       instructions="Choose the rank in column A and the role profile in column C.")

_section("WT", "DFX", "Fan Zone - Disruptive Effects")
for k, (nm, merc) in enumerate([("Disruptive Effects Officer", "DEO Officer"),
                                ("Disruptive Effects Operational Commander", "DEO Commander"),
                                ("Disruptive Effects Tactical Commander", "DEO Commander")]):
    _r("WT", "DFX", k * 2 + 1, f"Force Supplied - {nm}", "FS", "OFFICER", "Disruptive Effects",
       "Fan Zone", ranks=SP, mercury=merc)
    _r("WT", "DFX", k * 2 + 2, f"Mutual Aid Requested - {nm}", "MA", "OFFICER", "Disruptive Effects",
       "Fan Zone", ranks=SP, mercury=merc)

_section("WT", "FZA", "Fan Zone Armed Policing")
_r("WT", "FZA", 1, "Force Supplied - Fan Zone Armed Response (ARVs)", "FS", "OFFICER",
   "Armed Policing", "Fan Zone", ranks=SP, basis="Shifts", mercury="Firearms - ARV",
   instructions="Number of shifts")
_r("WT", "FZA", 2, "Mutual Aid Requested - Fan Zone Armed Response (ARVs)", "MA", "OFFICER",
   "Armed Policing", "Fan Zone", ranks=SP, basis="Shifts", mercury="Firearms - ARV",
   instructions="Number of shifts")

_section("WT", "WTP", "Wider Tournament Policing")
_r("WT", "WTP", 1, "Force Supplied - Control Room", "FS", "OFFICER", "Control Room", "General",
   ranks=SP, basis="Shifts", mercury="Control Room", instructions="Number of shifts")
_r("WT", "WTP", 2, "Force Supplied - Command Support", "FS", "OFFICER", "Command", "General",
   ranks=["CI", "INSP", "SGT", "PC"], basis="Shifts", mercury="Command Support",
   instructions="Number of shifts")
_r("WT", "WTP", 3, "Mutual Aid Requested - Command Support", "MA", "OFFICER", "Command", "General",
   ranks=["CI", "INSP", "SGT", "PC"], basis="Shifts", mercury="Command Support",
   instructions="Number of shifts")
_r("WT", "WTP", 4, "Force Supplied - Custody", "FS", "OFFICER", "Custody", "General",
   ranks=SP, basis="Shifts", mercury="Custody", instructions="Number of shifts")

_section("WT", "ADD", "Additional Resources")
_r("WT", "ADD", 1, "Force Supplied - Additional Resource", "FS", "ADDITIONAL",
   "Additional", "General", n_lines=5, rank_list="Ranks_All",
   instructions="Describe the resource in column B and choose the rank in column C.")
_r("WT", "ADD", 2, "Mutual Aid Requested - Additional Resource", "MA", "ADDITIONAL",
   "Additional", "General", n_lines=5, rank_list="Ranks_All",
   instructions="Describe the resource in column B and choose the rank in column C.")


RESOURCES = _RES
RESOURCE_BY_ID = {r["resource_id"]: r for r in RESOURCES}


def lines(res):
    """Yield one dict per collection line of a resource."""
    rt = res["row_type"]
    if rt in ("PSU_COUNT", "MA_PSU"):
        for i, lvl in enumerate(res["units"], 1):
            yield dict(line=i, rank=None, unit_level=lvl,
                       unit_list="Units_" + lvl, label="Level 2 PSU" if lvl == "L2" else "Level 3 BDU")
    elif res["ranks"]:
        for i, rk in enumerate(res["ranks"], 1):
            yield dict(line=i, rank=rk, unit_level=None, unit_list=None, label=RANK_LABEL[rk])
    else:
        for i in range(1, res["n_lines"] + 1):
            yield dict(line=i, rank=None, unit_level=None, unit_list=None, label="")


def row_key(rid, line):
    return f"{rid}|{line}"


# ---------------------------------------------------------------------------
# N/A rules (drive hatching, locking and data-quality checks)
# ---------------------------------------------------------------------------
def applicable(res, measure, day=None):
    """True if a (measure, day) cell is collectable for this resource."""
    rt = res["row_type"]
    if res["collection"] == "MD" and res["matchday_only"] and day not in (None, "MD"):
        return False
    if res["collection"] == "MD" and res["day_group"] == "NONMD" and day == "MD":
        return False                     # non-matchday resources: Match Day not used (legacy)
    if rt in ("PSU_COUNT", "MA_PSU"):
        return measure == "NUM"          # hours come from PSU config / MA package
    if res["supply"] == "MA" and measure == "TOD":
        # Mirrors the new template's hatching of Normal ToD on mutual-aid rows.
        # OPEN ITEM OI-01: switch to True if MA Normal ToD is to be collected.
        return MA_TOD_COLLECTED
    return True


MA_TOD_COLLECTED = False
NA_RULES_TEXT = [
    ("NA-01", "Matchday-only resources", "Only the Match Day column is collected; MD-2, MD-1 and MD+1 are N/A.",
     "Legacy Demand Sheet only used the Match Day figures for these rows."),
    ("NA-02", "PSU / serial count rows (force and mutual aid)", "Only 'Number required' is collected. Hours are N/A.",
     "Force-supplied hours come from the matching PSU configuration rows; mutual aid is costed by package."),
    ("NA-04", "Non-matchday public order (PSU configuration and counts)", "Match Day column is N/A.",
     "Legacy Demand Sheet never used the Match Day figures for non-matchday rows (they cover MD-2, MD-1, MD+1)."),
    ("NA-03", "Mutual aid officer rows", "Normal ToD is N/A (Extra Hours collected).",
     "Mirrors the hatching in the new draft template. OPEN ITEM OI-01."),
]
