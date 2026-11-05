# v1 template + original calculations (lift and shift)

`output/Euro_2028_Resource_Collection_v1_with_Calculations.xlsx` is the **new template v1, unchanged**
(Guidance, Match Day Resource, Data Info, Wider Tournament Resource – logos, banner, drop-downs and
formatting byte-for-byte as designed) with the **original workbook's calculation pages added behind it**,
using the original layout, formula style and logic, re-pointed at the new rows.

| Page | Visible | What it does (original logic) |
|---|---|---|
| Demand Sheet | hidden | Calendar 7 Jun – 10 Jul 2028, blocks Low / Medium / High / Total. **Row 4: type the fixture code under the risk column of its date.** For each risk column one relative day is used: standard rows MD > MD-1 > MD-2 > MD+1; non-matchday PSUs MD-1 > MD-2 > MD+1; matchday-only PSUs MD. PSU officers = officers per PSU × PSUs. |
| Cost Sheet | hidden | Demand × Cost Calculator for the same relative day. Fixture codes are read from the Demand Sheet (no second entry). |
| MD Demand Template | yes | Officers per resource per day (all risk levels) + total. |
| MD Cost Template | yes | Cost per resource per day (all risk levels) + total. |
| Fleet | hidden | Officers per resource per day. |
| Cost Calculator | hidden | Per officer: Extra Hours × hourly rate (Baseline Costs column X = Jun-2028). Mutual-aid PSU: package per unit per day. |
| WT Demand Sheet / WT Cost Sheet | hidden | Per typical wider-tournament day by risk level. |
| Wider Tournament Template | yes | Cost per day × days at each risk level + total. |
| Command Team & Planning, Other Resources, Baseline Costs, GDP Deflators | yes | Copied cell for cell from the original (force inputs blanked). |

## Rates (as the original)
Force supplied: Sergeant / Constable at host-force overtime (X31 / X32); Inspector and above £0, except
In-Stadium PSU Inspectors at the MA Inspector rate (X13). Mutual aid: X10–X15 by rank. Mutual-aid PSU /
BDU: X17 per unit per day. Additional Resources: MA Sergeant rate (X14).

New rows that did not exist in the original, using the nearest original rule (please confirm):
* **Rank (Drop Down) rows** (Mounted, Dogs, Roads Policing): the rate of the rank chosen (lookup table added
  on Baseline Costs B36:D39, same rate cells as above).
* **Role Profile rows** (Additional Public Order, Armed Response Overlay, CBRN Officer(s), WT Additional
  POPS): no rank is collected, so they use the original Additional Resources rule (MA Sergeant rate).
* **Wider Tournament days** (7 / 11 / 11 / 3) are set once on Baseline Costs C43:C46 instead of on every row.

## Proof it is the original logic
`validate_against_original.py` takes a completed original workbook (WMP), enters its force figures and
fixture codes into the new template, recalculates, and compares with the original's own results:

| Compared | Cells | Differences |
|---|---|---|
| Match Day demand (122 resource lines × 34 days × 3 risk levels) | 12,444 | 0 |
| Match Day cost | 12,444 | 0 * |
| Wider Tournament demand / cost (42 lines × 4 risk levels) | 168 / 168 | 0 / 0 |
| Wider Tournament Template totals | 42 | 0 |

\* With the original's fixture codes as entered, 88 cost cells differ (days 11–14 of the calendar, all £0 in
the original). The original asked for fixture codes twice – on the Demand Sheet and again on the Cost
Sheet – and they disagreed for one fixture (High on the Demand Sheet, Medium on the Cost Sheet), so that
fixture's costs were lost. Correcting the original's code gives **0 differences**. The new Cost Sheet reads
the codes from the Demand Sheet, so this cannot recur.

Also corrected (original defects, not logic): Mutual Aid ARV Overlay Sergeant / Constable referenced the
wrong rows / `#REF!` in the original; the lower half of the original Fleet page was `#REF!` and is not carried.

## Rebuild
`python3 euro2028/v1_lift_shift/build.py` (regenerates from `source/` and `legacy_static_pages.json`).
