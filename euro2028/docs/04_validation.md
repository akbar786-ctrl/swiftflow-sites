# Validation

## 1. Against the legacy workbook (re-keyed legacy return, legacy 2025 fixtures)

| Check | Compared | Result |
|---|---|---|
| Demand and cost by row, date and risk – officer rows (33 resources) | 9,384 cells × officers and cost | **identical** |
| Demand and cost – PSU / Public Order rows (configuration × count, mutual aid) | 3,264 cells × officers and cost | identical except 9 officer and 7 cost values, all on 13/16/19 June – the legacy halving defect (rules §5) |
| Command Team & Planning (legacy posts) | total | **£338,678.48 = legacy** |
| Other Resources (legacy hours) | total | **£75,426.98 = legacy** |
| Hourly rates and MA package at Jun-2028 / Jun-2027 | 14 values | **identical to 6 d.p.** |

The legacy return itself is not stored in the repository (OFFICIAL force data).

## 2. Backup workbook = Power BI rules
`tests/test_calc_workbook.py` fills the backup workbook with the synthetic return (including its
deliberate errors), recalculates it, and compares every MD line × date (officers and cost), every
WT line × risk, and the Summary totals with the reference engine: **all identical**.

## 3. Reconciliation targets – synthetic sample
Use after building the Power BI file. Source `output/samples/SYNTHETIC_Submission_TEST_FORCE.xlsx`.

**RuleSet = LEGACY (default)**

| Measure | Low | Medium | High | Super High |
|---|---|---|---|---|
| MD Officers per Fixture (MD-2..MD+1) | 686 | 924 | 1,147 | 1,491 |
| Officer Cost, MD | 92,140.16 | 123,948.67 | 153,468.24 | 198,066.25 |
| MA Package Cost, MD | 357,161.61 | 500,026.26 | 714,323.23 | 821,471.71 |
| WT Daily Officers | 197 | 260 | 322 | 416 |
| WT Daily Cost | 90,967.51 | 168,579.99 | 246,460.56 | 255,673.16 |
| WT Tournament Cost | 636,772.59 | 1,854,379.87 | 2,711,066.13 | 767,019.48 |

With the six synthetic fixtures in `RefFixture` (12 Jun High, 15 Jun Medium, 18 Jun High,
21 Jun Medium, 26 Jun Super High, 1 Jul Low – host force SYNTHETIC TEST FORCE):
Calendar Officers **6,319**, Calendar Cost **£4,452,372.51**, WT Tournament Cost **£5,969,238.07**.
These equal the backup workbook Summary (its Match Day count of 6,420 also includes 101 mutual-aid units).
Data quality: **2 errors, 4 warnings**.

**RuleSet = AGREED** (for reference): MD officer cost per fixture 382,066.10 / 514,269.72 / 635,398.47 /
822,886.79; WT tournament total £10,378,826.01; calendar cost £6,825,892.27.

## 4. Automated tests (26)
`python3 -m unittest discover -s euro2028/tests -v` – rates, layout, N/A rules, both rule sets,
PSU expansion, packages, calendar priority, data-quality codes, mapping completeness,
Command Team, Other Resources, backup workbook reconciliation.
