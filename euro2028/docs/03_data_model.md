# Power BI data model

## 1. Architecture

```
Force workbooks (folder) ──► Submissions ──► FactInputCell ──► FactInput ──┬─► FactDeployed ──► FactCalendar*
                                     │                                       └─► FactMAPackage
E28_Reference_Tables.xlsx ──► Dim* / Ref* tables (fnRef)                     DataQuality ◄── all of the above
                                                                              (* only when RefFixture is populated)
```

Star schema. Facts at the lowest useful grain; dimensions shared by MD and WT.

## 2. Tables

| Table | Grain | Source query |
|---|---|---|
| Submissions | one per force workbook (IsCurrent = latest) | `03_Submissions.pq` |
| FactInputCell | one per input cell (value state) | `04a_FactInputCell.pq` |
| FactInput | submission × line × risk × day (NUM, TOD, XH) | `04b_FactInput.pq` |
| FactDeployed | submission × resource × rank × risk × day × source | `05_FactDeployed.pq` |
| FactMAPackage | submission × MA PSU line × risk | `06_FactMAPackage.pq` |
| FactCalendar | fixture × date × deployed row | `09_FactCalendar.pq` |
| DataQuality | one per issue | `08_DataQuality.pq` |
| DimResource, DimResourceLine, DimRank, DimRelativeDay, DimRoleProfile, DimForce, MapOldToNew, LogDecision, LogOpenItem, RefCapacity, RefEquipmentRatio, RefScenarioUplift, RefFixture, RefFixturePlan | reference | `fnRef("<name>")` |
| DimRisk | risk + WT days | `07c_DimRisk.pq` |
| DimRate | supply × rank effective rates | `07a_DimRate.pq` |
| DimMAPackage | unit type package rate | `07b_DimMAPackage.pq` |
| DimDate | calendar 1 Sep 2026 – 31 Mar 2029 | DAX: `DimDate = CALENDAR(DATE(2026,9,1), DATE(2029,3,31))` |
| DimMercurySkill | distinct skills | DAX: `DimMercurySkill = DISTINCT(UNION(VALUES(DimResource[MercurySkill]), VALUES(RefCapacity[MercurySkill])))` |

## 3. Relationships (all single direction, many-to-one)

| From (many) | To (one) |
|---|---|
| FactDeployed[ResourceID] | DimResource[ResourceID] |
| FactDeployed[RankCode] | DimRank[RankCode] |
| FactDeployed[RiskCode] | DimRisk[RiskCode] |
| FactDeployed[DayCode] | DimRelativeDay[DayCode] |
| FactDeployed[ForceKey] | DimForce[Force] |
| FactDeployed[RateKey] | DimRate[RateKey] |
| FactDeployed[SubmissionID] | Submissions[SubmissionID] |
| FactMAPackage[ResourceID / RiskCode / DayCode / ForceKey] | DimResource / DimRisk / DimRelativeDay / DimForce |
| FactInput[ResourceID / RiskCode / DayCode / ForceKey] | DimResource / DimRisk / DimRelativeDay / DimForce |
| FactInputCell[SubmissionID] | Submissions |
| FactCalendar[Date] | DimDate[Date] |
| FactCalendar[ResourceID / RankCode / RiskCode / RateKey / HostForce] | DimResource / DimRank / DimRisk / DimRate / DimForce |
| RefFixture[MatchDate] | DimDate[Date] (inactive; used by `Fixtures in Range`) |
| DataQuality[SubmissionID] | Submissions[SubmissionID] |
| RefCapacity[Force / RankCode / MercurySkill] | DimForce / DimRank / DimMercurySkill |
| DimResource[MercurySkill] | DimMercurySkill[MercurySkill] |

## 4. Power Query vs DAX

| Power Query (refresh time) | DAX (query time) |
|---|---|
| Reading files by key row / key column | Quantities and costs (`measures.dax`) |
| Keeping only the latest submission per force | Per-fixture, planned and calendar views |
| Blank / zero / N/A / invalid rules | WT tournament totals (× national days) |
| PSU expansion (count × configuration) | Capacity gap, utilisation |
| MA package counting (basis parameter) | Equipment, scenario uplifts |
| Effective rates (base × uplift, zero rules) | DQ counts, completion % |
| Calendar expansion with additive overlap | |
| All data-quality checks | |

Costs are measures (not columns) so that rate or parameter changes need only a
refresh, and every cost responds to every slicer.

## 5. Build steps in Power BI Desktop

1. Create parameters `pSubmissionsFolder` and `pReferenceFile` (`00_Parameters.pq`).
2. Create `fnRef` and `fnReadCollection` as blank queries (paste the function code).
3. Create the reference queries listed in `01_fnRef.pq`.
4. Create queries 03 → 09 in order, pasting each file. Disable load on
   `RefParameter`, `RefBaseRate`, `RefUpliftStep`, `RefZeroCost`, `RefMAPackage`.
5. Create `DimDate` and `DimMercurySkill` (DAX, above) and the relationships in §3.
6. Create a `_Measures` table and paste `measures.dax`.
7. Point `pSubmissionsFolder` at a folder containing
   `output/samples/SYNTHETIC_Submission_TEST_FORCE.xlsx`; refresh; check the
   reconciliation figures in `04_validation.md`.

Performance: `fnReadCollection` is called by `FactInputCell` and `DataQuality`.
For 48 forces (~800k cells) this is acceptable; if refresh is slow, make a
`RawCells` staging query (the `Cells` step of 04a) and reference it from both.

## 6. Report pages (mapped to the question prompts)

| Page | Answers |
|---|---|
| Schedule | venues, fixtures, kick-off, same-day venues, calendar (from `RefFixture` once set) |
| Resourcing – Match Day | per Mercury skill: resources on MD-2 / MD-1 / MD / MD+1 by risk |
| Resourcing – Wider Tournament | per Mercury skill: WT resources by risk |
| Daily deployment | resources per day in a date range (calendar; per-fixture until fixtures set) |
| Budgeting | cost by skill, fan zone, night-time economy (Town Centre / Wider Community), armed, roads, drones, CBRN, MA vs force supplied – force and national |
| Equipment | protected carriers, horses, C-UAS kit, LFR vans (`Equipment Required`, OI-05) |
| Scenarios | Op Brock, Op Stack, Israel qualifies, national reserve (OI-06) |
| Capacity | demand peak vs capacity by force / region / skill (OI-07) |
| Data quality | submissions received, errors / warnings by force, section completion |
| Reference | decisions, open items, old-to-new mapping |
