# EURO 2028 – Resource Collection and Power BI Model

Forces complete one collection template. Two routes turn returns into the dashboard, and both give **identical results** using the legacy calculation rules:

| Route | How | When |
|---|---|---|
| **A – Power BI** (primary) | Power BI reads the completed collection templates and does every calculation in Power Query / DAX. | Normal use |
| **B – Backup workbook** (fallback) | Paste a return into the backup calculation workbook; it calculates everything in Excel (same pages as the old workbook); Power BI reads its output pages. | If Route A is unavailable |

> OFFICIAL. No force data is stored here. `output/samples/` holds **synthetic** test files only.

## Files

| File | Purpose |
|---|---|
| `output/E28_Collection_Template_v2.0.0.xlsx` | Sent to forces. Same form as the draft (columns A–C; risk level on top; Super High on Match Day). |
| `output/E28_Backup_Calculation_Workbook_v2.0.0.xlsx` | Route B. Same input pages + the legacy calculation pages (Demand Sheet, Cost Calculator, Cost Sheet, MD/WT Demand & Cost Templates, Fleet, Command Team & Planning, Other Resources, Baseline Costs, GDP Deflators) + Summary. |
| `output/E28_Reference_Tables.xlsx` | Everything Power BI looks up: resources, ranks, rates, uplifts, parameters, mapping, decisions, open items. |
| `powerbi/` | Power Query (one query per file, numbered in build order) and `measures.dax`. |
| `docs/` | [Power BI build guide](docs/03_powerbi_build_guide.md) · [Calculation rules](docs/02_calculation_rules.md) · [Validation](docs/04_validation.md) · [Audit](docs/01_audit.md) |

## Evidence the numbers are right

* Re-keyed legacy return: **12,648 legacy Demand Sheet / Cost Sheet cells compared (officers and cost); all identical except 9 officer and 7 cost values** – all on 13, 16 and 19 June, caused by the legacy halving defect on those three dates only (audit C3).
* Command Team & Planning **£338,678.48** and Other Resources **£75,426.98** – identical to the legacy workbook.
* All hourly rates and the £35,716.16 package identical to the legacy Baseline Costs (to 6 d.p.).
* Backup workbook (Excel formulas) = Power BI rules (Python reference engine) on every line, date and risk level.
* 26 automated tests: `python3 -m unittest discover -s euro2028/tests -v`

## Changing things

Edit `build/catalogue.py` (resources) or `build/reference_data.py` (rates, parameters), then
`python3 euro2028/build/build_all.py` and re-run the tests. Resource IDs are permanent keys.
The rule set is one parameter: `RuleSet` = `LEGACY` (default) or `AGREED`.
