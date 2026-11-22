# Template v3 – planned changes (noted 22 Nov 2026, awaiting the national team's current template)

## 1. Guidance sheet (first tab)
Purpose · who completes it and by when · how to submit (file name = force name, SharePoint) ·
colour key (white = type, yellow = choose, grey/hatched = locked) · step-by-step by section ·
worked examples, incl. PSU / Serial configuration (e.g. 1 Inspector, 3 Sergeants, 21 Constables
per PSU, then the number of PSUs) · "do not miss" checklist (officers AND Normal ToD AND Extra Hours
on every line used – no hours = no cost) · common mistakes · definitions (MD-2…MD+1, risk levels,
FS / MA, ToD) · FAQs · contact · version / OFFICIAL. Include the team's past example screenshots.

## 2. Locked cells
Whole sheet protected; only input cells unlocked. Locked: headers, labels, totals, hidden Power BI
header row and Section column, hatched N/A cells. List of cells to lock to come from the team.
Password held by NPoCC (protection prevents accidents, it is not security). Power BI reads
protected sheets normally.

## 3. Drop-downs
In-cell lists (no hidden list sheet): PSU / Serial configuration and counts 1–25 (range to
confirm), origin, rank on Additional Resources. Error alert "Stop" so only list values are accepted.

## 4. Completeness (missing ToD / Extra Hours = missing cost)
Validation cannot force a cell to be filled, so:
* conditional formatting – if Officers is entered and Normal ToD / Extra Hours is blank on that
  day, the blank cells turn red; hours with no officers also flag;
* a status column after the totals per row ("Missing hours", "Missing rank", OK);
* a completion panel per sheet: number of incomplete lines, "READY TO SUBMIT" only at 0;
* forces enter 0 rather than leaving blank when there genuinely are no extra hours;
* PSU check: PSUs entered but configuration blank → flag;
* Power BI Data Quality page catches anything that still gets through.
Macros (e.g. block saving until complete) are not used – VBA is blocked on force laptops.

## 5. Placement rules (keep Power BI stable)
Do not move existing rows/columns. New items go in the banner area, after the totals columns, or
on the Guidance sheet. Power BI selects columns by the hidden header names, so extra columns at
the end do not break it.

## Open questions
Rows/cells to lock · password · 1–25 for which cells · mutual aid Normal ToD (currently hatched)
required or not · submission deadline and contact for Guidance · example screenshots.
