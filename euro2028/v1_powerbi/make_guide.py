"""Writes output/EURO_2028_Power_BI_Step_by_Step.txt - the guide plus every query and measure, in order."""
import glob
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "output", "EURO_2028_Power_BI_Step_by_Step.txt")

GUIDE = r"""EURO 2028 - FROM THE FORCES' TEMPLATE TO A POWER BI DASHBOARD                       OFFICIAL
Step-by-step guide
==========================================================================================

THE SHORT ANSWER TO "DO I NEED A MASTER TABLE?"
  Yes - but you never build it by hand. Power BI builds it every time you press Refresh:
    * Power Query opens every force's template, reads ONLY "Match Day Resource" and
      "Wider Tournament Resource", and turns the grid (rows = resources, columns = risk level /
      day / measure) into one long MASTER table: one row per force, resource, rank, risk level
      and day, with officers, hours and cost.
    * To know what each Excel row means, it uses a ROW MAP (in the reference workbook):
      "row 120 = Football Spotting DFO, Force supplied, Sergeant, use the FS Sergeant rate".
      The template's rows never move, so the map never changes.
  You do NOT copy, reshape or merge anything by hand, and forces' files are never edited.

------------------------------------------------------------------------------------------
PART A - THE FILES
------------------------------------------------------------------------------------------
  EURO_2028_Resource_Collection_Template.xlsx
      Send to forces. Looks exactly like the original, with risk level on row 3 (Low |
      Medium | High | Super High) and day/measure on row 4. It contains two hidden named
      ranges (MD_Data, WT_Data) that Power BI reads - forces never see them.
  EURO_2028_Power_BI_Reference.xlsx
      Keep yourself (NOT in the returns folder). Row maps, column maps, hourly rates, mutual
      aid PSU package, Wider Tournament days, fixture list. Edit only the yellow cells.
  EURO_2028_Power_BI_TEST_PACK.zip
      Three made-up returns, a reference file with test fixtures, and the figures Power BI
      must show. Use it to prove your build before real returns arrive.

------------------------------------------------------------------------------------------
PART B - BEFORE YOU START (once)
------------------------------------------------------------------------------------------
  1. Make a folder for returns, e.g.  C:\EURO2028\Returns\
  2. Put EURO_2028_Power_BI_Reference.xlsx somewhere else, e.g.  C:\EURO2028\
  3. Ask forces to return the template unchanged (no inserted/deleted rows or columns,
     no renamed sheets) and SAVE EACH RETURN AS THE FORCE NAME, e.g.
         West Midlands Police.xlsx
     The file name becomes the force name in Power BI and must match the fixture list.
  4. To test first: unzip the TEST PACK, put the three "Sample Force" files in the returns
     folder and use EURO_2028_Power_BI_Reference_SAMPLE.xlsx as the reference file.

------------------------------------------------------------------------------------------
PART C - CREATE THE QUERIES (Power BI Desktop)
------------------------------------------------------------------------------------------
  5. Open Power BI Desktop > Home > Transform data. The Power Query Editor opens.
  6. For EACH query at the end of this file, in the order given:
       a. Home > New Source > Blank Query
       b. Home > Advanced Editor > delete everything > paste the query text
          (from the line after "===== QUERY" down to the next "=====" line) > Done
       c. In the Query Settings pane (right), Name = the name shown in the ===== line.
       d. If marked "do not load": right-click the query > untick "Enable load".
     The two parameters (ReturnsFolder, ReferenceFile): change the path in quotes to yours.
     Ignore errors on early queries - they clear once the later ones exist.
  7. Click MD_Fact. You should see one row per force / resource / rank / risk / day with
     Officers and costs. That is your master table for Match Day. WT_Fact is the same for
     the Wider Tournament; Calendar_Fact puts Match Day on real dates.
  8. Home > Close & Apply.

------------------------------------------------------------------------------------------
PART D - MODEL (relationships, sorting, measures)
------------------------------------------------------------------------------------------
  9. Model view: drag column onto column to create these (many-to-one, single direction):
       MD_Fact[Force] -> DimForce[Force]          MD_Fact[Risk] -> DimRisk[Risk]
       MD_Fact[Day]   -> DimDay[Day]
       WT_Fact[Force] -> DimForce[Force]          WT_Fact[Risk] -> DimRisk[Risk]
       Calendar_Fact[Force] -> DimForce[Force]    Calendar_Fact[Risk] -> DimRisk[Risk]
       Calendar_Fact[Day]   -> DimDay[Day]        Calendar_Fact[Calendar date] -> DimDate[Date]
 10. Sorting: select DimRisk[Risk] > Column tools > Sort by column > Risk order.
     Do the same for DimDay[Day] and DimDay[Day name] (sort by Day order).
 11. Measures: Home > Enter data > name it "Measures" > Load. Select the Measures table >
     New measure > paste ONE measure line (from the DAX section at the end) > Enter.
     Repeat for each. Format the cost measures: Measure tools > Currency (GBP), 0 decimals.

------------------------------------------------------------------------------------------
PART E - CHECK IT IS RIGHT (with the TEST PACK)
------------------------------------------------------------------------------------------
 12. Report view: add a Table visual with DimForce[Force], DimRisk[Risk] and the measures
     MD Officer Deployments, MD PSUs Serials, MD Overtime Cost, MD Mutual Aid PSU Cost,
     MD Cost (Original Basis). Compare with EXPECTED_Power_BI_results_SAMPLE.xlsx -
     every number should match. Do the same for the WT and Calendar blocks and the cards.
 13. Open the DataQuality table: it lists anything to chase (rank not chosen, PSUs given
     without officers-per-PSU, old template, force on the fixture list with no return).
     Empty = clean.
 14. When it matches: point ReturnsFolder / ReferenceFile at the real folder and file
     (Home > Transform data > Edit parameters) and Refresh.

------------------------------------------------------------------------------------------
PART F - BUILD THE DASHBOARD
------------------------------------------------------------------------------------------
 15. OVERVIEW page
       Cards: Forces Returned, Tournament Cost (Original Basis), Calendar Cost (Original
       Basis), WT Cost (Original Basis), Issues to Chase.
       Stacked bar: Axis DimForce[Force]; values Calendar Cost (Original Basis) and
       WT Cost (Original Basis).
       Column chart: Axis DimRisk[Risk]; value MD Officer Deployments.
 16. MATCH DAY page (what ONE fixture needs)
       Slicers: DimRisk[Risk] (single select), DimForce[Force], MD_Fact[Supplied by],
       MD_Fact[Rank], MD_Fact[Role / Skill].
       Cards: MD Officer Deployments, MD Peak Day Officers, MD PSUs Serials,
       MD Cost (Original Basis), MD Mutual Aid Share.
       Bar: MD_Fact[Section] by MD Officer Deployments.  Column: DimDay[Day name] by
       MD Officer Deployments.  Matrix: rows DimForce[Force], columns DimRisk[Risk],
       values MD Cost (Original Basis).
 17. FIXTURE CALENDAR page (needs fixtures in the reference file)
       Stacked column: Axis DimDate[Date]; Legend DimForce[Force]; value Calendar Officer
       Deployments. Second chart the same with Calendar Cost (Original Basis).
       Matrix: rows DimForce[Force], columns DimDate[Date], values Calendar Officer
       Deployments. Tooltip: Calendar_Fact[Fixtures].
 18. WIDER TOURNAMENT page
       Cards: WT Officer Days (Tournament), WT Cost (Original Basis), WT Cost incl Normal Duty.
       Stacked bar: Axis DimForce[Force]; Legend DimRisk[Risk]; value WT Cost (Original Basis).
       Matrix: rows DimForce[Force], columns DimRisk[Risk], values WT Daily Officers.
 19. RESOURCE BREAKDOWN page
       Matrix: rows MD_Fact[Section] > MD_Fact[Resource] > MD_Fact[Rank]; columns
       DimRisk[Risk]; values MD Officer Deployments and MD Cost (Original Basis).
       Slicers: Supplied by, Line type.
 20. DATA QUALITY page: Table visual of DataQuality.

------------------------------------------------------------------------------------------
PART G - EACH COLLECTION ROUND
------------------------------------------------------------------------------------------
 21. Save each force's return into the returns folder as "<Force name>.xlsx".
 22. Update fixtures / rates in the reference file if needed (yellow cells).
 23. Power BI > Home > Refresh. Check the Data Quality page, chase forces, refresh again.

------------------------------------------------------------------------------------------
HOW THE CALCULATIONS WORK (the original collection workbook's rules, proven exactly)
------------------------------------------------------------------------------------------
  * Officers on an officer row = "Number required". Hours are per officer.
  * Cost of an officer row = Number x Extra Hours x hourly rate. Rate = Rates table for the
    row's "Rate supplied by" + "Rate rank" (row map); rank from column C where the row has a
    rank drop-down (Mounted, Dogs, Roads Policing).
  * Force supplied Inspector and above are not costed (rate 0). In Stadium PSU Inspectors use
    the mutual aid Inspector rate. Role / Additional rows use the mutual aid Sergeant rate.
  * PSUs / Serials: force supplied officers = number of PSUs x officers per PSU (the PSU /
    Serial configuration rows at the top), costed with the configuration rows' Extra Hours.
    Mutual aid PSUs = number of PSUs x package per PSU per day (Settings).
  * Normal ToD cost is shown separately ("incl Normal Duty" measures); "Original Basis"
    leaves it out, as the original did.
  * Calendar: MD-2..MD+1 fall on the days around each fixture. If a force has two fixture-days
    on one date at one risk level, ONE is used: Match Day > MD-1 > MD-2 > MD+1 (non-matchday
    PSUs MD-1 > MD-2 > MD+1; matchday-only PSUs Match Day only).
  * Wider Tournament total = typical-day figure x days at that risk level (7 / 11 / 11 / 3).
  * Proof: the original collection workbook's own figures, entered in this template, give the
    same officers and cost on every date and the same Wider Tournament total - 0 differences
    (also with the figures moved to Super High).

TO CONFIRM (rules carried over from the original - change the Rates table if they change)
  * Force supplied Inspectors and above at GBP 0.
  * Role / Additional rows (e.g. Armed Response Overlay, CBRN officers) at the mutual aid
    Sergeant rate - the template has no rank for them.
  * Mutual aid package applies to Level 3 BDUs / Serials as well as Level 2 PSUs.

TROUBLESHOOTING
  * "No readable EURO 2028 templates": the folder path is wrong or empty.
  * A force is missing: see Data Quality - usually an old template or a renamed sheet.
  * Numbers look shifted: someone inserted rows/columns - ask for the return on a fresh
    template (Data Quality flags "Old layout" when row 3 is not Low/Medium/High/Super High).
  * Fixture force not found: the fixture list force name must equal the file name.
==========================================================================================
"""


def main(out=OUT):
    parts = [GUIDE]
    for f in sorted(glob.glob(os.path.join(HERE, "powerbi", "*.pq"))):
        text = open(f).read().strip()
        first = text.split("\n")[0]
        name = re.search(r"Query name:\s*(\w+)", first).group(1)
        note = re.search(r"\(([^)]*)\)", first).group(1)
        parts.append(f"\n===== QUERY: {name}   ({note}) {'=' * max(3, 60 - len(name) - len(note))}\n{text}\n")
    parts.append("\n===== DAX MEASURES " + "=" * 60 + "\n" + open(os.path.join(HERE, "powerbi", "13_measures.dax")).read())
    with open(out, "w", newline="\r\n") as fh:
        fh.write("".join(parts))
    return out


if __name__ == "__main__":
    print(main())
