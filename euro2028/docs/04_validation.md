# Stage 4 – Validation strategy and results

## 1. Approach

Three independent layers, so no result relies on one implementation:

1. **Rule tests** – `tests/test_engine.py` fills copies of the real template
   exactly as a force would (including mistakes) and checks the engine against
   hand calculations that use the **legacy workbook's cached rates**.
2. **Power BI reconciliation** – after the first refresh of the Power BI file on
   the synthetic sample, the measures must match the engine figures in §3.
3. **Legacy comparison** – re-key a legacy force return into the new template and
   compare against the legacy outputs, explaining every difference by a
   decision or finding ID (§4). Differences are expected; none may be unexplained.

A formula is never assumed correct because it exists in the old workbook: every
legacy behaviour is either adopted by a recorded decision or replaced.

## 2. Rule tests (23 tests, all passing)

Run: `python3 -m unittest discover -s euro2028/tests -v`

| Area | Test | Proves |
|---|---|---|
| Rates | `test_hourly_overtime_rates_reproduce_legacy` | All 12 HF/MA June-2028 rates equal `Baseline Costs!X10:X15, X27:X32` to 6 d.p. |
| | `test_package_reproduces_legacy` | MA PSU package = £35,716.1613 (`Baseline Costs!X17`) |
| | `test_force_supplied_inspector_and_above_zero` | D-05 / OI-02 zero rules; detective ranks use uniform rates |
| | `test_mutual_aid_inspector_charged`, `test_missing_rate_returns_none` | OI-02 default; missing rate is null, not 0 |
| Template | `test_every_row_key_once`, `test_column_keys` | Template matches the catalogue; 48 MD and 12 WT input columns; risk-first order |
| | `test_na_cells_locked_inputs_unlocked` | N/A rules applied and protected |
| Demand | `test_psu_expansion_and_cost` | 6 L2 PSUs × (1/3/18) = 6/18/108 officers, 1,848 hours (guidance worked example); cost by rank |
| | `test_psu_count_without_config_flags` | DQ13 |
| | `test_calendar_overlap_adds` | D-02: MD+1 of one fixture + MD-2 of the next are added |
| Mutual aid | `test_ma_package_bases` | Matchday package per fixture; non-matchday peak-day default; legacy per-day alternative |
| | `test_serial_package_has_no_rate` | OI-03 surfaces as DQ14 |
| | `test_mutual_aid_officer_extra_hours_only`, `test_ma_tod_switch` | NA-03 / OI-01 switch |
| Pricing | `test_role_rows_priced_by_selected_rank`, `test_additional_resource` | D-07 |
| WT | `test_wider_tournament_days` | per-day cost × national days (D-08) |
| Data quality | `test_invalid_values`, `test_blank_zero_and_section_status`, `test_missing_force_flagged`, `test_latest_submission_wins` | DQ01, DQ04, DQ05, DQ12, DQ16; blank vs zero |
| Mapping | `test_mapping_complete_and_valid` | every new resource is mapped from a legacy row or marked Added; every legacy row maps or is Retired |

## 3. Power BI reconciliation targets (synthetic sample)

Source: `output/samples/SYNTHETIC_Submission_TEST_FORCE.xlsx` (synthetic test
values only). Engine defaults as in `RefParameter`.

| Check | L | M | H | SH |
|---|---|---|---|---|
| MD officer-days per fixture (MD-2..MD+1) | 728 | 984 | 1,211 | 1,572 |
| MD officer cost per fixture (£) | 418,884.01 | 566,340.58 | 692,688.34 | 895,294.88 |
| MD MA package cost per fixture (£) | 357,161.61 | 500,026.26 | 678,607.07 | 785,755.55 |
| WT officers per day | 187 | 247 | 304 | 394 |
| WT cost per day (£) | 112,392.91 | 147,959.22 | 182,582.51 | 236,660.99 |
| WT MA package cost per day (£) | 71,432.32 | 142,864.65 | 214,296.97 | 214,296.97 |
| WT tournament cost (£, × 7 / 11 / 11 / 3 days) | 1,286,776.62 | 3,199,062.53 | 4,365,674.25 | 1,352,873.86 |

Also: WT tournament total £10,204,387.26; PSU-expanded officers 3,492; MD MA
packages 65; FactInput 5,072 rows; FactDeployed 945 rows; FactMAPackage 44 rows.
Data quality: exactly 6 issues – DQ05 (1), DQ06 (1), DQ07 (1), DQ09 (1), DQ10 (1), DQ12 (1).

Measures to compare: `MD Officers per Fixture`, `Officer Cost` and
`MA Package Cost` filtered to MD by `DimRisk`; `WT Daily Officers`,
`WT Daily Cost`, `WT Tournament Cost`; `DQ Errors` + `DQ Warnings`.

## 4. Legacy comparison

The legacy cached outputs (MD £18,794,476.42; 10,584 officer-days; WT
£1,039,983.75) come from one legacy return. To compare like with like:

1. Re-key that return into the new template (Level 1 PSU lines have no new
   equivalent – record them separately).
2. Run the model in **legacy-replication settings**: `MA_ToD_Collected` = FALSE,
   ToD rate set to 0 (A1), `MAPSU_Basis_NonMatchday` = `PER_UNIT_PER_DAY`,
   `RefFixture` = the five 2025 legacy fixtures with their row-4 risk letters.
3. Explain the residual difference by these known causes:

| Cause | Direction | Finding |
|---|---|---|
| Overlapping days added instead of first-match-wins | new higher | C4 / D-02 |
| MD+1 PSU halving on three dates not applied | new higher | C3 |
| Cost Sheet copy errors (BS8:BS15, D40:F48) removed | either | C5 |
| Additional resources priced by rank | either | C6 / D-07 |
| In-Stadium Inspectors now £0 | new lower | A5 / D-05 |
| Level 1 PSU retired | new lower | D-09 |
| Normal ToD costed (when switched back on) | new higher | D-03 |

4. Sign off the comparison table with the project team before go-live.

## 5. Ongoing checks after each refresh

* `DQ Errors` = 0 for every force included in national figures.
* `Check: Direct Officer Hours` equals the sum of the template's *Total hours* columns.
* `Forces Submitted` matches the submissions log.
* Any change to `RefParameter`, rates or the catalogue is followed by re-running the unit tests.
