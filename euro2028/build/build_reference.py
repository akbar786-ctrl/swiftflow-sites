"""
Builds E28_Reference_Tables.xlsx - every reference table the Power BI model
loads, as named Excel tables (one per sheet). Generated from catalogue.py,
reference_data.py and legacy_mapping.py.
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, os.path.dirname(__file__))
import catalogue as C  # noqa: E402
import legacy_mapping as M  # noqa: E402
import reference_data as R  # noqa: E402

TABLES = []  # (sheet, table, description, headers, rows)


def table(sheet, name, desc, headers, rows):
    TABLES.append((sheet, name, desc, headers, [list(r) for r in rows]))


def _tables():
    res_rows, line_rows = [], []
    for r in C.RESOURCES:
        cfg = r["psu_config"] or {}
        res_rows.append([r["resource_id"], r["collection"], r["section_code"],
                         C.SECTIONS[(r["collection"], r["section_code"])], r["name"], r["supply"],
                         dict(C.SUPPLY_TYPES)[r["supply"]], r["row_type"], r["portfolio"], r["location"],
                         r["matchday_only"], r["day_group"], r["legacy_insp_ma"], r["basis"], r["mercury"],
                         "DRAFT - confirm (OI-04)", cfg.get("L2"), cfg.get("L3"), r["instructions"]])
        for ln in C.lines(r):
            line_rows.append([C.row_key(r["resource_id"], ln["line"]), r["resource_id"], ln["line"],
                              ln["rank"], C.RANK_LABEL.get(ln["rank"]) if ln["rank"] else None,
                              ln["unit_level"], ln["unit_list"] or r["rank_list"], r["role_list"], ln["label"]])
    table("DimResource", "DimResource", "One row per collection resource (Resource ID is the stable key).",
          ["ResourceID", "Collection", "SectionCode", "Section", "Resource", "SupplyCode", "Supply", "RowType",
           "Portfolio", "Location", "MatchdayOnly", "DayGroup", "LegacyInspMA", "Basis", "MercurySkill", "MercuryStatus",
           "PsuConfigL2", "PsuConfigL3", "Instructions"], res_rows)
    table("DimResourceLine", "DimResourceLine", "One row per collection line (Row Key = ResourceID|Line).",
          ["RowKey", "ResourceID", "Line", "FixedRankCode", "FixedRank", "UnitLevel", "RankOrUnitList",
           "RoleList", "LineLabel"], line_rows)
    table("DimRank", "DimRank", "Ranks; RateRank is used to look up rates (detective ranks map to uniform).",
          ["RankCode", "Rank", "RateRank", "Sort", "IsDetective"], C.RANKS)
    table("DimRisk", "DimRisk", "Risk levels.", ["RiskCode", "Risk", "Sort"], C.RISK_LEVELS)
    table("DimRelativeDay", "DimRelativeDay", "Relative days. WT = typical wider-tournament day.",
          ["DayCode", "Day", "OffsetDays", "Sort"], C.MD_DAYS + [("WT", "Wider Tournament day", None, 5)])
    table("DimMeasure", "DimMeasure", "Collected measures.", ["MeasureCode", "Measure"], C.MEASURES)
    table("DimRoleProfile", "DimRoleProfile", "Role profiles (new template Data Info sheet).",
          ["Portfolio", "RoleProfile"], C.ROLE_PROFILES)
    table("DimUnitType", "DimUnitType", "PSU / serial unit types.",
          ["UnitCode", "UnitType", "PsuLevel", "IsSerial"], C.UNIT_TYPES)
    table("DimForce", "DimForce", "Forces and region (region DRAFT - OI-07).",
          ["Force", "Region", "RegionStatus"], [(f, rg, "DRAFT - confirm") for f, rg in R.FORCES])
    table("RefBaseRate", "RefBaseRate", "Hourly rates 2025-26 before uplift. HF = host force, MA = mutual aid.",
          ["RateTable", "RankCode", "RateType", "BaseRate", "Source"],
          [(t, rk, "TOD", tod, st) for t, rk, tod, ot, st, so in R.BASE_RATES] +
          [(t, rk, "OT", ot, so) for t, rk, tod, ot, st, so in R.BASE_RATES])
    table("RefUpliftStep", "RefUpliftStep", "Cumulative uplift factors (legacy method replicated).",
          ["RateTable", "EffectiveFrom", "Factor", "Method"], R.UPLIFT_STEPS)
    table("RefZeroCost", "RefZeroCost", "Supply/rank combinations costed at GBP 0.",
          ["SupplyCode", "RankCode", "AppliesTo", "Basis"], R.ZERO_COST_RULES)
    table("RefMAPackage", "RefMAPackage", "Mutual-aid PSU package, 2025-26 base (uplifted by RateTable MAPSU).",
          ["UnitType", "BasePackage", "Note"], R.MA_PSU_PACKAGE)
    table("RefParameter", "RefParameter", "Model parameters and switches.",
          ["Parameter", "Value", "Type", "Description", "Source"],
          [(n, v.isoformat() if isinstance(v, dt.date) else v, t, d, s) for n, v, t, d, s in R.PARAMETERS])
    table("RefWTDays", "RefWTDays", "National assumption: wider-tournament days at each risk level (D-08).",
          ["RiskCode", "Days", "Source"], R.WT_DAY_ASSUMPTIONS)
    table("RefGDPDeflator", "RefGDPDeflator", "GDP deflators (ONS / OBR, March 2026).",
          ["FinancialYear", "Index", "PctChange"], R.GDP_DEFLATORS)
    table("RefPayAward", "RefPayAward", "Police pay award assumptions.", ["FinancialYear", "PayAward"], R.PAY_AWARDS)
    table("RefAnnualCost", "RefAnnualCost", "Annual employable cost 2025-26 (MA, HF, CTP = Command Team).",
          ["CostTable", "RankCode", "AnnualCost"], R.ANNUAL_COSTS)
    names = {r["resource_id"]: r["name"] for r in C.RESOURCES}
    map_rows = []
    for coll, row, old, new_ids, ctype, note in M.MAPPING:
        for nid in (new_ids or [None]):
            map_rows.append((coll, row, old, nid, names.get(nid), ctype, note))
    for nid, note in M.ADDED.items():
        map_rows.append((nid[:2], None, None, nid, names[nid], "Added", note))
    table("MapOldToNew", "MapOldToNew", "Legacy survey resource -> new Resource ID.",
          ["Collection", "LegacySurveyRow", "LegacyResource", "ResourceID", "Resource", "ChangeType", "Note"],
          map_rows)
    table("RefNARule", "RefNARule", "Not-applicable rules driving hatching/locking.",
          ["RuleID", "AppliesTo", "Rule", "Reason"], C.NA_RULES_TEXT)
    table("LogDecision", "LogDecision", "Decisions recorded with the project team.",
          ["DecisionID", "Decision"], R.DECISIONS)
    table("LogOpenItem", "LogOpenItem", "Open items requiring confirmation.",
          ["ItemID", "Topic", "Detail", "Status"], R.OPEN_ITEMS)
    # Structures awaiting data (one blank row each)
    table("RefEquipmentRatio", "RefEquipmentRatio", "TO BE COMPLETED (OI-05): equipment per resource.",
          ["Equipment", "ResourceID", "MercurySkill", "Basis", "Ratio", "Source", "Status"], [[None] * 7])
    table("RefScenarioUplift", "RefScenarioUplift",
          "TO BE COMPLETED (OI-06): Op Brock / Op Stack / Israel qualifies / National reserve.",
          ["Scenario", "Force", "Portfolio", "MercurySkill", "RankCode", "RiskCode", "DayCode",
           "UpliftType", "Value", "Source"], [[None] * 10])
    table("RefCapacity", "RefCapacity", "TO BE COMPLETED (OI-07): available officers by force / skill / rank.",
          ["Force", "Portfolio", "MercurySkill", "RankCode", "AvailableOfficers", "AsAtDate", "Source"],
          [[None] * 7])
    table("RefFixture", "RefFixture",
          "TO BE COMPLETED when fixtures are set. Leave empty: the model then reports per fixture.",
          ["FixtureID", "MatchDate", "KickOff", "Venue", "HostForce", "Stage", "TeamA", "TeamB", "RiskCode"],
          [[None] * 9])
    table("RefFixturePlan", "RefFixturePlan",
          "OPTIONAL what-if: planned number of fixtures per host force per risk level.",
          ["Force", "RiskCode", "Fixtures", "Note"], [[None] * 4])


def build(path):
    _tables()
    wb = Workbook()
    ws = wb.active
    ws.title = "README"
    ws["A1"] = "EURO 2028 Reference Tables"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A2"] = ("Generated by euro2028/build/build_reference.py - edit the Python source, not this file, "
                "except the 'TO BE COMPLETED' tables which the project team maintains.")
    ws["A4"], ws["B4"], ws["C4"] = "Table", "Rows", "Description"
    for c in "ABC":
        ws[f"{c}4"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["C"].width = 110
    for i, (sheet, name, desc, headers, rows) in enumerate(TABLES, 5):
        ws.cell(i, 1, name)
        ws.cell(i, 2, len([r for r in rows if any(v is not None for v in r)]))
        ws.cell(i, 3, desc)
        sh = wb.create_sheet(sheet[:31])
        for ci, h in enumerate(headers, 1):
            sh.cell(1, ci, h)
            sh.column_dimensions[L(ci)].width = max(12, min(60, len(h) + 4))
        for ri, row in enumerate(rows, 2):
            for ci, v in enumerate(row, 1):
                cell = sh.cell(ri, ci, v)
                if isinstance(v, dt.date):
                    cell.number_format = "yyyy-mm-dd"
                if isinstance(v, str) and len(v) > 60:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
        t = Table(displayName=name, ref=f"A1:{L(len(headers))}{max(2, len(rows) + 1)}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        sh.add_table(t)
        sh.freeze_panes = "A2"
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "..", "output",
                                                             "E28_Reference_Tables.xlsx")
    print(build(os.path.abspath(out)))
