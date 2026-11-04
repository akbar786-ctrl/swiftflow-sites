# Calculation specification

This is the single statement of every business rule. The Python engine
(`engine/e28_engine.py`) and the Power BI model (`powerbi/`) both implement
exactly these rules; the unit tests in `tests/` prove the engine does.

## 1. What forces submit

One workbook per force, completed for **one typical fixture** (Match Day) and
**one typical day** (Wider Tournament), at **every** risk level
(Low, Medium, High, Super High). Fixtures are not yet known (D-01).

| Dimension | Match Day (MD) | Wider Tournament (WT) |
|---|---|---|
| Risk level | L, M, H, SH | L, M, H, SH |
| Relative day | MD-2, MD-1, MD, MD+1 | per typical day |
| Measures | Number required, Normal ToD (hrs per officer), Extra Hours (hrs per officer) | same |

Layout: risk level is the top header, relative day beneath, measure beneath
that (D-11). The model reads the hidden key row, so column order is irrelevant.

### Row types

| Row type | Number means | Hours | Costed by |
|---|---|---|---|
| PSU_CONFIG | officers of that rank in ONE PSU / serial | per officer | via PSU_COUNT |
| PSU_COUNT | force-supplied PSUs / serials | N/A (from config) | officers × hours × rate |
| MA_PSU | mutual-aid PSUs / serials | N/A | package |
| OFFICER | officers (or shifts where stated) of the fixed rank | per officer | hours × rate |
| RANK_SELECT | officers of the rank chosen | per officer | hours × rate |
| ROLE | officers of the role profile and rank chosen | per officer | hours × rate (by rank, D-07) |
| ADDITIONAL | officers described by the force, with supply and rank chosen | per officer | hours × rate (by rank, D-07) |

### Blank, zero, not applicable, invalid

| State | Meaning | Treatment |
|---|---|---|
| BLANK | not answered | null; reported via section status |
| BLANK_AS_ZERO | blank in a section marked *Complete* / *Not required* | 0 |
| VALUE | valid number | used |
| NA | hatched, locked cell (rules `RefNARule`) | ignored; a value typed in is excluded and reported (DQ06) |
| INVALID | text, negative, decimal Number, hours > 24 | excluded and reported (DQ04 / DQ05) |

N/A rules: matchday-only resources collect Match Day only (NA-01); PSU count rows
collect Number only (NA-02); mutual-aid Normal ToD is N/A (NA-03, OI-01).

## 2. Demand

* **Direct rows:** Officers = Number; ToD hours = Number × Normal ToD; Extra hours = Number × Extra Hours.
* **PSU expansion:** for each force-supplied count line (Level 2 or Level 3) at a risk and day,
  officers of each rank = units × officers-per-unit from the linked configuration
  at the **same risk and day**; hours = officers × that configuration's hours.
  A count with an empty configuration is reported (DQ13) and not expanded.
* **MD+1 PSU factor:** 1 (legacy halving not adopted; parameter `MD1_PSU_Factor`).
* **Overlapping dates (D-02):** when fixtures are known, each fixture's demand is
  placed on fixture date + offset; coinciding dates are **added**.

## 3. Rates

`rate = base 2025-26 rate × Π uplift factors effective on or before RateEffectiveDate (1 Jun 2028)`

| Supply | Table | Normal ToD rate | Extra Hours rate |
|---|---|---|---|
| Force supplied | Host Force (`Baseline Costs!B18:H23`) | Hourly Employable Cost | Hourly Overtime Cost |
| Mutual aid | Mutual Aid (`Baseline Costs!B1:H6`) | Hourly Employable Cost | Hourly Overtime Cost |

Uplift steps (legacy method, D-04): MA Sep-26 × (GDP deflator index 26/27 + pay award 26/27),
Sep-27 × 1.035, Sep-28 × 1.03. HF Sep-26 × 1.04, Sep-27 × 1.035, Sep-28 × 1.03.
Detective ranks use the equivalent uniform rate.

Zero-cost rules: force-supplied Inspector (D-05) and, as legacy, force-supplied
Chief Inspector and above (OI-02); ACC has no rate in either table.

Resulting June-2028 rates (match legacy column X to 6 d.p.):

| Rank | HF ToD | HF OT | MA ToD = OT |
|---|---|---|---|
| Constable | 53.20 | 45.97 | 54.22 |
| Sergeant | 66.19 | 55.21 | 67.46 |
| Inspector | 0 (FS) | 0 (FS) | 67.40 |
| Chief Inspector | 0 (FS) | 0 (FS) | 74.90 |
| Superintendent | 0 (FS) | 0 (FS) | 92.38 |
| Chief Superintendent | 0 (FS) | 0 (FS) | 106.65 |

## 4. Cost

* **Officer cost** = ToD hours × ToD rate + Extra hours × OT rate.
* **Mutual-aid PSU package** = packages × £35,716.16 (Jun-2028; base £31,601.07 + 5% admin, uplifted).
  * Matchday-only MA PSUs: one package per unit per fixture (D-06).
  * Non-matchday MA PSUs: one package per unit at the peak day count across MD-2..MD+1 (OI-03 default;
    `PER_UNIT_PER_DAY` reproduces legacy).
  * WT MA PSUs: one package per unit per typical day (legacy; OI-03).
  * Serials: no rate yet (OI-03) – reported as DQ14, not costed.
* **Match Day outputs** are per fixture by risk level. With `RefFixturePlan` (fixtures per force per risk) the
  model also gives a planned total; with `RefFixture` it gives a dated calendar.
* **Wider Tournament total** = Σ over risk levels of (cost per typical day × national days at that risk:
  L 7, M 11, H 11, SH 3 – D-08).

## 5. Decisions

| ID | Decision |
|---|---|
| D-01 | Risk per fixture is an assumption; forces complete every risk level; model works per fixture until fixtures are set. |
| D-02 | Overlapping relative days are added; MD and MD+1 are separate categories. |
| D-03 | Normal ToD costed for all officers. |
| D-04 | Host Force and Mutual Aid rate tables are final. |
| D-05 | Inspectors £0 for Normal ToD and Extra Hours. |
| D-06 | £35,716 package covers travel day in, match day and travel day out. |
| D-07 | Additional resources and role profiles priced by selected rank. |
| D-08 | WT keeps Super High; WT days per risk are national assumptions. |
| D-09 | Level 1 PSU, Servator Complex / Non-Complex, CTSFO Overlay, MA Command retired. |
| D-10 | Command Team & Planning, Other Resources and Fleet brought into Power BI. |
| D-11 | Risk-first header layout; Super High added to Match Day. |

## 6. Open items (model default in brackets)

| ID | Item | Switch |
|---|---|---|
| OI-01 | Mutual-aid Normal ToD – template hatches it, D-03 says cost all ToD. (Not collected.) | `MA_ToD_Collected`; `catalogue.MA_TOD_COLLECTED` |
| OI-02 | Force-supplied CI and above £0; mutual-aid Inspector+ charged at MA rate. (As legacy.) | `RefZeroCost` |
| OI-03 | "Mutual aid costs are different": packages for non-matchday MA PSUs, Level 3 BDU price, serial price. | `MAPSU_Basis_*`, `RefMAPackage` |
| OI-04 | Mercury skill mapping (draft). | `DimResource[MercurySkill]` |
| OI-05 | Equipment ratios (carriers, horses, C-UAS kit, LFR vans). | `RefEquipmentRatio` |
| OI-06 | Op Brock / Op Stack / Israel / national reserve uplifts. | `RefScenarioUplift` |
| OI-07 | Force regions and capacity. | `DimForce`, `RefCapacity` |
| OI-08 | Protest Removal as dedicated rows vs role profiles (double-count risk). | Guidance |
