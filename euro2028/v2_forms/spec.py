"""
Content of the two-page EURO 2028 collection form (Match Day + Wider Tournament).

Every row a force can fill in is defined here once; the form builder, the
Power BI reference implementation and the tests all read this file.

Row fields
  section   group heading (Power BI: Section)
  resource  what the resource is (Power BI: Resource)
  supply    "Force supplied" | "Mutual aid" | None (= drop-down: force chooses)
  role      "-" (not applicable) | name of a ROLE_LISTS list (drop-down)
  rank      a fixed rank | "Number of PSUs / Serials" | RANK_DD (drop-down)
  days      Match Day only: "ALL" | "MD_ONLY" | "NON_MD"   (which day columns may be used)
  free      True when the force types the resource name (Additional Resources)
  resource_list  optional drop-down list for the Resource cell (WT additional PO)

Source: rows, ranks, drop-down lists and Matchday / Non-Matchday rules of the
BLANK v1 template (Match Day Resource / Wider Tournament Resource / Data Info).
Changes from v1: Role / Skill and Rank are separate columns; Mounted, Dogs and
Roads Policing list Inspector / Sergeant / Constable rows instead of a rank
drop-down; PSU / Serial counts and officers are entered on one line each.
"""

FS, MA = "Force supplied", "Mutual aid"
UNITS = "Number of PSUs / Serials"
RANK_DD = "<rank>"
NA = "-"

RISKS = ["Low", "Medium", "High", "Super High"]
DAYS = ["MD-2", "MD-1", "MD", "MD+1"]
DAY_NAMES = {"MD-2": "Match Day -2", "MD-1": "Match Day -1", "MD": "Match Day", "MD+1": "Match Day +1"}
MEASURES = ["Officers", "Normal hrs", "Extra hrs"]

RANKS = ["Chief Superintendent", "Superintendent", "Chief Inspector", "Inspector", "Sergeant", "Constable"]
ROLE_LISTS = {
    "PSU_L2": ["Level 2 PSU", "Level 2 Serial"],
    "PSU_L3": ["Level 3 BDU", "Level 3 Serial"],
    "PUBLIC_ORDER": ["Forward Intelligence Team Officer (FIT)", "Police Liaison Officer (PLT)", "Medic",
                     "PO Bronze Commander", "PO Safety Advisor", "PO Loggist"],
    "PUBLIC_ORDER_WT": ["Forward Intelligence Team Officer (FIT)", "Police Liaison Officer (PLT)", "Medic",
                        "PO Bronze Commander", "PO Safety Advisor", "PO Loggist", "Non-Complex Protest Removal",
                        "Working at Height - Protest Removal", "De-Bonding - Protest Removal"],
    "ARMED": ["Firearms Tactical Advisor", "Specialist Firearms Officer", "Specialist Strategic Firearms Officer",
              "Specialist Tactical Firearms Officer", "Strategic Firearms Commander", "ARV", "CTSFO"],
    "CBRN": ["CBRN - Enhanced Responder", "CBRN - Operational Commander (Bronze)", "CBRN - Search",
             "CBRN - Tactical Advisor", "CBRN - Tactical Commander (Silver)"],
}
RESOURCE_LISTS = {
    "WT_ADDITIONAL_PO": ["Additional Public Order - Wider Community", "Additional Public Order - Fan Zone"],
}
ISC = ["Inspector", "Sergeant", "Constable"]
SC = ["Sergeant", "Constable"]


def row(section, resource, supply, rank, role=NA, days="ALL", free=False, resource_list=None):
    return dict(section=section, resource=resource, supply=supply, role=role, rank=rank, days=days, free=free,
                resource_list=resource_list)


def both(section, resource, ranks, days="ALL", fs_only=False):
    out = []
    for sup in ([FS] if fs_only else [FS, MA]):
        out += [row(section, resource, sup, rk, days=days) for rk in ranks]
    return out


def psu_block(section, resource, days):
    """PSU / Serial type: force supplied = number of units + officers by rank; mutual aid = number of units."""
    out = []
    for level, lst in (("Level 2", "PSU_L2"), ("Level 3", "PSU_L3")):
        name = f"{resource} - {level}"
        out.append(row(section, name, FS, UNITS, role=lst, days=days))
        out += [row(section, name, FS, rk, days=days) for rk in ISC]
        out.append(row(section, name, MA, UNITS, role=lst, days=days))
    return out


def role_rows(section, resource, role_list, n, days="ALL", resource_list=None):
    out = []
    for sup in (FS, MA):
        out += [row(section, resource, sup, RANK_DD, role=role_list, days=days, resource_list=resource_list)
                for _ in range(n)]
    return out


def additional(section, n):
    return [row(section, "", None, RANK_DD, free=True) for _ in range(n)]


# ------------------------------------------------------------------ Match Day
PO = "Public Order / Public Safety"
MD_ROWS = (
    psu_block(PO, "Public Order (Non-Matchday)", "NON_MD")
    + psu_block(PO, "Town Centre Public Order Policing (Matchday Only)", "MD_ONLY")
    + psu_block(PO, "Stadium Vicinity Public Order Policing (Matchday Only)", "MD_ONLY")
    + psu_block(PO, "In Stadium Public Order Policing (Matchday Only)", "MD_ONLY")
    + both("Evidence Gathering", "Evidence Gathering Team Officer", SC)
    + both("Protest Removal", "Working at Height Complex Protest Removal", ISC)
    + both("Protest Removal", "Non-Complex Protest Removal", ISC)
    + both("Protest Removal", "Debonding Protest Removal", ISC)
    + role_rows("Additional Public Order / Public Safety", "Additional Public Order", "PUBLIC_ORDER", 6)
    + both("Mounted", "Mounted Commander", ISC)
    + both("Mounted", "Mounted Officer", ISC)
    + both("Dogs", "Dogs GDP", ISC)
    + both("Dogs", "Dogs Public Order Serial Commander", ISC)
    + both("Dogs", "Dogs Public Order", ISC)
    + both("Dogs", "Dogs Explosives", ISC)
    + both("Football Policing", "Football Spotting DFO", SC)
    + both("Football Policing", "Football Spotting OFO", SC)
    + both("Roads Policing", "Roads Policing Officer", ISC)
    + both("Roads Policing", "Roads Policing Officer - Motorcyclists", ISC)
    + both("Roads Policing", "Roads Policing Officer - VIP Motorcyclist Escort", ISC)
    + both("Armed Policing", "Operational Firearms Commander", ISC)
    + both("Armed Policing", "Armed Response (ARVs)", ISC)
    + both("Armed Policing", "CTSFOs", ISC)
    + both("Armed Policing", "AFOs", ISC)
    + both("Armed Policing", "Rifle Officer", SC)
    + both("Armed Policing", "Protection Officer", SC)
    + role_rows("Additional Armed Policing", "Armed Response Overlay", "ARMED", 7)
    + both("Disruptive Effects", "Servator Disruptive Effects", SC)
    + both("Disruptive Effects", "Servator Disruptive Effects Operational Commander", SC)
    + both("Disruptive Effects", "Servator Disruptive Effects Tactical Commander", SC)
    + both("Search", "LSOs", ISC)
    + both("Search", "PolSAs", ISC)
    + both("Search", "CT Secco", ISC)
    + both("Unmanned Aerial Systems (Drones)", "Counter UAS Bronze Commander", ISC)
    + both("Unmanned Aerial Systems (Drones)", "Counter UAS Effector Operator", ISC)
    + both("Unmanned Aerial Systems (Drones)", "Counter UAS Observers", ISC)
    + both("Unmanned Aerial Systems (Drones)", "Counter UAS Tactical Advisor", ISC)
    + both("Unmanned Aerial Systems (Drones)", "Drone", ISC)
    + both("Live Facial Recognition", "LFR Operational Commander", ISC)
    + both("Live Facial Recognition", "LFR Operator", ISC)
    + both("CBRN", "General Responder", ISC)
    + role_rows("CBRN", "CBRN Officers", "CBRN", 5)
    + both("Force Command (non Public Order)", "Command",
           ["Assistant Chief Constable"] + RANKS, fs_only=True)
    + both("Control Room", "Control Room", ISC, fs_only=True)
    + both("Custody", "Custody", SC, fs_only=True)
    + both("Crime", "Crime", ["Detective Inspector", "Detective Sergeant", "Detective Constable"])
    + both("Intelligence", "Intel", ISC)
    + both("Ancillary", "Ancillary Team", SC)
    + additional("Additional Resources", 9)
)

# ------------------------------------------------------------ Wider Tournament
WT_ROWS = (
    psu_block(PO, "Wider Community Public Order Policing", "ALL")
    + psu_block(PO, "Fan Zone Public Order Policing", "ALL")
    + both(PO, "Operationally Deployable non-Public Order Policing", ISC)
    + both(PO, "Evidence Gatherers", SC)
    + role_rows("Additional Public Order / Public Safety", "", "PUBLIC_ORDER_WT", 9,
                resource_list="WT_ADDITIONAL_PO")
    + both("Fan Zone - Disruptive Effects", "Disruptive Effects Officer", SC)
    + both("Fan Zone - Disruptive Effects", "Disruptive Effects Operational Commander", SC)
    + both("Fan Zone - Disruptive Effects", "Disruptive Effects Tactical Commander", SC)
    + both("Fan Zone - Armed Policing", "Fan Zone Armed Response (ARVs)", SC)
    + both("Wider Tournament Policing", "Control Room", SC, fs_only=True)
    + both("Wider Tournament Policing", "Command Support", ["Chief Inspector"] + ISC)
    + both("Wider Tournament Policing", "Custody", SC, fs_only=True)
    + additional("Additional Resources", 9)
)


def md_headers():
    return [f"{r} | {d} | {m}" for r in RISKS for d in DAYS for m in MEASURES]


def wt_headers():
    return [f"{r} | {m}" for r in RISKS for m in MEASURES]


KEY_COLS = ["Section", "Resource", "Supplied by", "Role / Skill", "Rank"]
