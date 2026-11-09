"""
Bundles the Power BI set-up into ONE text file: the steps, every Power Query query (in the
order to create them) and the DAX measures. Usage: python3 make_powerbi_pack.py [out.txt]
"""
import glob
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output", "EURO_2028_Power_BI_Setup.txt")

STEPS = r"""EURO 2028 - POWER BI SET-UP                                                      OFFICIAL
================================================================================

WHAT YOU NEED
  1. A folder holding the forces' completed forms (EURO_2028_Resource_Collection_Form.xlsx,
     one file per force, any file names).
  2. EURO_2028_Reference_Data.xlsx (rates, mutual aid PSU package, Wider Tournament days,
     fixture list). Keep it OUTSIDE the returns folder.

TO TEST FIRST: put the three files from samples\returns in a folder, use
EURO_2028_Reference_Data_SAMPLE.xlsx as the reference file, and compare your visuals with
EXPECTED_Power_BI_results_SAMPLE.xlsx. Every figure should match.

STEP 1 - CREATE THE QUERIES
  Power BI Desktop > Home > Transform data (opens Power Query).
  For EACH query below, in order:
    Home > New Source > Blank Query, then Home > Advanced Editor,
    delete what is there, paste the query text (everything between the ===== lines),
    click Done, then rename the query (right-click > Rename) to the name shown.
  For the two parameters (ReturnsFolder, ReferenceFile) change the path in quotes to your own.
  Queries marked "do not load": right-click > untick "Enable load".

STEP 2 - CLOSE & APPLY
  Home > Close & Apply. Then check the DataQuality table: every row is something to chase
  with a force (blank rank, text typed in a number cell, a force on the fixture list with no
  return, ...). Empty = clean.

STEP 3 - RELATIONSHIPS (Model view, drag column onto column, all "many to one", single direction)
  MD_Fact[Force]         -> DimForce[Force]       MD_Fact[Risk]   -> DimRisk[Risk]
  MD_Fact[Day]           -> DimDay[Day]
  WT_Fact[Force]         -> DimForce[Force]       WT_Fact[Risk]   -> DimRisk[Risk]
  Calendar_Fact[Force]   -> DimForce[Force]       Calendar_Fact[Risk] -> DimRisk[Risk]
  Calendar_Fact[Day]     -> DimDay[Day]           Calendar_Fact[Calendar date] -> DimDate[Date]
  Sort columns: DimRisk[Risk] by [Risk order]; DimDay[Day] and [Day name] by [Day order]
  (select the column > Column tools > Sort by column).

STEP 4 - MEASURES
  Home > Enter data > name the table "Measures" > Load. Then for each measure at the end of
  this file: select the Measures table > New measure > paste ONE line (name = formula).
  Format the cost measures as Currency (GBP), 0 decimals.

STEP 5 - SUGGESTED PAGES
  Overview      cards: Forces Returned, Tournament Cost (Original Basis), MD Cost (Original Basis)
                (pick a Risk on a slicer), WT Cost (Original Basis), Issues to Chase.
  Match Day     matrix rows DimForce[Force], columns DimRisk[Risk], values MD Cost (Original Basis);
                bar chart MD Officer Deployments by MD_Fact[Section]; slicers Risk, Supplied by, Rank,
                Role / Skill.
  Calendar      line/column chart DimDate[Date] x Calendar Officer Deployments and
                Calendar Cost (Original Basis); matrix Force x Date.
  Wider Tourn.  matrix Force x Risk: WT Daily Officers, WT Cost (Original Basis).
  Data quality  table visual of DataQuality.

HOW THE CALCULATIONS WORK (same rules as the original collection workbook)
  * Officers on a line = number of officers; hours are per officer.
  * Overtime cost = Officers x Extra hrs x hourly rate for the line's Supplied by + Rank
    (Rates table; a Rate overrides line wins for a specific resource).
  * Force supplied Inspector and above are not costed (rate 0) - original rule; change the
    Rates table if that should change. In Stadium PSU Inspectors use the mutual aid
    Inspector rate (Rate overrides) - original rule.
  * Mutual aid PSUs / Serials: number of units x package per unit per day (Settings).
  * Normal duty cost (Officers x Normal hrs x rate) is shown separately; "Original Basis"
    totals leave it out, "incl Normal Duty" totals include it.
  * MD figures are for ONE fixture at each risk level. The Calendar uses the fixture list:
    MD-2..MD+1 land on the dates around each fixture; when a force has two fixture-days on
    the same date and risk level, one is used: Match Day > MD-1 > MD-2 > MD+1 (non-matchday
    PSU lines: MD-1 > MD-2 > MD+1).
  * Wider Tournament total = typical-day figure x days at that risk level (WT days table).
  * "Officer deployments" counts officers x days.

================================================================================
"""


def main(out):
    files = sorted(glob.glob(os.path.join(HERE, "powerbi", "*.pq")))
    parts = [STEPS]
    for f in files:
        text = open(f).read().strip()
        name = re.search(r"Query name:\s*(\w+)", text).group(1)
        note = "do not load" if "do not load" in text.split("\n")[0] or "function" in text.split("\n")[0] \
            else ("parameter" if "parameter" in text.split("\n")[0] else "LOAD")
        parts.append(f"===== QUERY: {name}   ({note}) " + "=" * max(0, 50 - len(name) - len(note)) + "\n")
        parts.append(text + "\n\n")
    parts.append("===== DAX MEASURES " + "=" * 60 + "\n")
    parts.append(open(os.path.join(HERE, "powerbi", "13_measures.dax")).read())
    with open(out, "w", newline="\r\n") as fh:
        fh.write("".join(parts))
    return out


if __name__ == "__main__":
    print(main(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
