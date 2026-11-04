# Calculation rules

One statement of every rule. Implemented identically by Power BI (`powerbi/`), the
backup calculation workbook and the Python reference engine (`engine/`).
`RuleSet = LEGACY` (default) reproduces the old workbook; `AGREED` applies the revised rules.

## 1. What forces enter
For **one fixture** (Match Day) and **one typical day** (Wider Tournament), at every risk
level (Low, Medium, High, Super High): Number required, Normal ToD hours and Extra Hours
per officer, for MD-2, MD-1, Match Day and MD+1.

| Row | Number means | Hours from |
|---|---|---|
| PSU configuration | officers of that rank in ONE PSU / serial | the row |
| Public Order (force) | PSUs / serials needed | the matching PSU configuration |
| Public Order (mutual aid) | PSUs / serials needed | – (package) |
| Officer rows (fixed rank, rank drop-down, role profile, additional) | officers (or shifts where column A says so) | the row |

Not applicable (hatched, locked): matchday-only rows – Match Day only; non-matchday rows – no
Match Day; PSU count rows – Number only; mutual aid – no Normal ToD.
Blank = not answered (0 if the section is marked Complete). Text, negatives, decimal Numbers
and hours above 24 are excluded and reported.

## 2. Demand
* Officers = Number. PSU officers = PSUs × officers per PSU for the same risk and day.
* **Calendar** (when fixtures are entered): each fixture places its MD-2 … MD+1 demand on the
  dates around it.

| | LEGACY | AGREED |
|---|---|---|
| A date is several relative days of same-risk fixtures | one only, first in priority: standard rows MD > MD-1 > MD-2 > MD+1; non-matchday rows MD-1 > MD-2 > MD+1; matchday-only rows MD | all added |
| Fixtures at different risk levels on nearby dates | added | added |

## 3. Cost
| | LEGACY | AGREED |
|---|---|---|
| Officer cost | officers × Extra Hours × overtime rate | officers × (ToD × hourly rate + Extra Hours × overtime rate) |
| Force-supplied Inspector and above | £0, except In-Stadium PSU Inspectors at the MA Inspector rate | £0 |
| Additional Resources | MD: MA Sergeant rate; WT: HF Constable rate | rate of the rank selected |
| Mutual-aid PSU / BDU | £35,716.16 per unit per day | per unit per fixture (non-matchday: peak day) |
| Wider Tournament total | cost per day × national days (L 7, M 11, H 11, SH 3) | same |

Rates: force supplied = Host Force table, mutual aid = Mutual Aid table (Baseline Costs), uplifted
to June 2028 by the legacy method. Detective ranks use the uniform rate.

| Rank | Force supplied (OT) | Mutual aid (OT) |
|---|---|---|
| Constable | £45.97 | £54.22 |
| Sergeant | £55.21 | £67.46 |
| Inspector | £0 | £67.40 |
| Chief Inspector | £0 | £74.90 |
| Superintendent | £0 | £92.38 |
| Chief Superintendent | £0 | £106.65 |

## 4. Other pages
* **Command Team & Planning:** posts × (annual salary / 12) × uplift (deflator 26/27 + pay award 26/27,
  then +3.5% Sep-27, +3% Sep-28), for every month from Start to End **inclusive**.
* **Other Resources:** FY 2027-28 hours × HF overtime rate at Jun-2027; FY 2028-29 hours × rate at Jun-2028.
* **Fleet:** officers per resource per date × items per officer (rounded up), where ratios are entered.

## 5. Known legacy defects not reproduced
| Defect | Effect |
|---|---|
| PSUs halved on MD+1 on three dates only (13, 16, 19 June, rows 7–17) | legacy lower on those dates |
| MD+1 branch missing in 12 cells (one date column, rows 7–15 and 105) | legacy zero in those cells |
| Cost Sheet copy errors (BS8:BS15, D40:F48) | legacy wrong rate in those cells |
| Command Team M23:N28 repeat previous month | no effect on values |

## 6. Decisions and open items
Decisions D-01 – D-12 and open items OI-01 – OI-08 are in `E28_Reference_Tables.xlsx`
(`LogDecision`, `LogOpenItem`). Open items each change one parameter or table entry.
