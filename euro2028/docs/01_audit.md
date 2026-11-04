# Stage 1 – Audit of the legacy workbook

Source: *Base Assumptions MPSTH – Uniform.xlsx* (16 sheets, ~70,000 formulas).
The audit read every sheet cell by cell, grouped formulas by relative pattern,
and cross-checked references. The original workbook was not modified.

## 1. Sheet inventory

| Sheet | State | Role |
|---|---|---|
| MD Resource Collection Survey | visible | Input: ~230 resource × rank rows; MD-2/MD-1/MD/MD+1 × Low/Medium/High × Number/Normal ToD/Extra Hours; SUMIF totals AO:AW. |
| WT Resource Collection Survey | visible | Input: Low/Medium/High/Super High × 3 measures. |
| Demand Sheet | hidden | Calendar 7 Jun–10 Jul (**2025 dates**), 4 cols/date (L, M, H, total). Row 4 = match letter. Hand-wired IF chains pick MD/MD-1/MD-2/MD+1 survey figures. |
| Cost Calculator | hidden | Survey **Extra Hours** × hourly rate, or £35,716 per MA PSU. |
| Cost Sheet | hidden | Demand × Cost Calculator with the same day logic. |
| MD Demand Template / MD Cost Template | visible | Outputs: sum of L+M+H per date. |
| WT Demand Template / WT Cost Sheet / WT Cost Template | mixed | WT chain; cost/day × days per risk (7/11/11/3). |
| Baseline Costs | visible | Mutual Aid and Host Force rate tables, uplifts, MA PSU package. |
| GDP Deflators | visible | ONS/OBR deflators (March 2026). |
| Command Team & Planning | visible | Monthly cost of planning posts. |
| Other Resources | visible | Other overtime hours by category. |
| Fleet | hidden | Broken (980 `#REF!`). |
| Sheet1 | hidden | PSU/Serial list. |

## 2. Findings

### Critical

| # | Finding | Evidence | Resolution |
|---|---|---|---|
| C1 | Risk level per match is encoded by **which sub-column the match letter is typed into**; each formula checks its own column only. Low produced 0 for the whole tournament. | `Demand Sheet!Z4, AK4, AX4, BI4, CD4` | Confirmed as intended (D-01). Replaced by an explicit risk field per fixture. |
| C2 | Calendar uses 2025 dates. | `Demand Sheet!D3:EF3` | Model works per fixture; calendar is generated from `RefFixture` when set. |
| C3 | MD+1 rule inconsistent: on 13/16/19 June, rows 7–17, MD+1 is **added and PSUs halved**; elsewhere MD+1 is exclusive and not halved. | 99 cells `Demand Sheet!AB7:AD17, AN, AZ` | D-02: add both; no halving (parameter `MD1_PSU_Factor` = 1). |
| C4 | Overlapping relative days: the IF chain keeps only the first match (MD > MD-1 > MD-2 > MD+1), dropping demand. | All Demand / Cost formulas | D-02: demands are added. |
| C5 | Copy errors: `Cost Sheet!BS8:BS15` test `BP5…BP12` instead of row 4 and use the Inspector rate `$N$7`; `Cost Sheet!D40:F48` reference wrong Cost Calculator rows. | Cell level | Eliminated (no cell-by-cell formulas). |
| C7 | MD+1 branch missing from the IF chain in 12 cells of one date column (rows 7–15 and 105). | Demand Sheet | Eliminated (generated formulas). |
| C6 | Additional Resources priced at a fixed rate regardless of role (MD: MA Sergeant; WT: arbitrary mix). | `Cost Calculator` rows 269–277, 366–381 | D-07: priced by the selected rank. |

### Costing assumptions (resolved with the project team)

| # | Legacy behaviour | Decision |
|---|---|---|
| A1 | Only Extra Hours costed. | D-03: Normal ToD costed for all officers (MA rows: see OI-01). |
| A2–A4 | Rate tables, uplift method (additive deflator + pay award for MA; pay award only for HF). | D-04: tables final; method replicated exactly. |
| A5 | Force-supplied Inspector+ £0, except In-Stadium PSU Inspectors charged at MA rate. | D-05: Inspectors £0 (In-Stadium exception removed). |
| A6 | £35,716 flat per MA PSU per day. | D-06: package = travel in + MD + travel out; basis per OI-03. |
| A7 | Three salary bases (MA, Host Force, Command Team). | Kept as separate tables (`RefAnnualCost`). |
| A8 | `Other Resources!G3` labelled FY 27/28 but uses Jun-2028 rates. | Template columns labelled FY 2027-28 / FY 2028-29. |

### Structural / hygiene

* WT survey totals AQ:AV contain `#REF!` and drifting header references.
* `Cost Sheet!E7` unique formula unlike neighbours (leftover test).
* `Command Team & Planning!M23:N28` reference the previous month's value instead of the rate row (latent).
* Used range extends to row 1,048,576 on four sheets; stray `WT Cost Sheet!DC1048576`.
* Labels unreliable as keys (typos, trailing / non-breaking spaces, "Sergeants" vs "Sergeant"); WT "Number required" means shifts on some rows.
* Data validation stored as an extension and lost outside Excel.

Verified clean: Demand rows reference the correct survey rows across all 102
date columns; scenario columns are consistent; all 215 Cost Sheet rows map to
label-matching Cost Calculator rows except the C5 cells.

Re-keying the legacy return into the new model reproduces 12,648 legacy Demand / Cost Sheet cells
exactly, apart from the C3 dates – see `04_validation.md`.

## 3. Legacy output baselines (cached values)

| Output | Value |
|---|---|
| MD Cost Template, grand total | £18,794,476.42 |
| MD Demand Template, officer-days | 10,584 |
| WT Cost Template, total | £1,039,983.75 |
| Other Resources, total | £75,426.98 |
| Demand Sheet by scenario (officer-deployments) | L 0 / M 3,273 / H 7,311 |

## 4. New draft template observations

* Hatching (N/A cells) appears to drift down rows across successive Normal ToD columns in the Match Day -2 and Match Day blocks. The rebuilt template generates N/A cells from rules (`RefNARule`).
* Mutual-aid Normal ToD is hatched, which conflicts with D-03 (OI-01).
* Role-profile rows had no rank column; the rebuilt template collects both.
* Rank list was Inspector/Sergeant/Constable only; role profiles such as Strategic Firearms Commander need higher ranks (`Ranks_All`).
* Protest Removal exists as dedicated rows and as role profiles (OI-08).
* Typo "CBRN – Renhanced Responder" corrected.

The old-to-new resource mapping is in `E28_Reference_Tables.xlsx` → `MapOldToNew`
(source: `build/legacy_mapping.py`).
