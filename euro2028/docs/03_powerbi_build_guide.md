# Power BI build guide (step by step)

Time: about half a day. You need Power BI Desktop, the reference workbook, and
a folder of returns (start with `output/samples/SYNTHETIC_Submission_TEST_FORCE.xlsx`).

## 1. Parameters
Home → Transform data → Manage Parameters → New (type **Text**):

| Name | Value |
|---|---|
| `pSubmissionsFolder` | folder holding completed collection templates |
| `pReferenceFile` | full path to `E28_Reference_Tables.xlsx` |
| `pBackupFolder` | Route B only: folder holding completed backup workbooks |

## 2. Queries
For each file, in order: New Source → Blank Query → Advanced Editor → paste the
file → rename the query to the name in the file header.

| Order | File | Query name | Load |
|---|---|---|---|
| 1 | `01_fnRef.pq` | `fnRef` | – |
| 2 | (in `01_fnRef.pq` list) | `DimResource`, `DimResourceLine`, `DimRank`, `DimRelativeDay`, `DimRoleProfile`, `DimForce`, `MapOldToNew`, `LogDecision`, `LogOpenItem`, `RefEquipmentRatio`, `RefScenarioUplift`, `RefCapacity`, `RefFixture`, `RefFixturePlan` = `fnRef("<name>")` | yes |
| 2 | (same) | `RefParameter`, `RefBaseRate`, `RefUpliftStep`, `RefZeroCost`, `RefMAPackage`, `RefAnnualCost` = `fnRef("<name>")` | **no** |
| 3 | `07a_DimRate.pq` | `DimRate` | yes |
| 4 | `07b_DimMAPackage.pq` | `DimMAPackage` | yes |
| 5 | `07c_DimRisk.pq` | `DimRisk` | yes |
| 6 | `02_fnReadCollection.pq` | `fnReadCollection` | – |
| 7 | `03_Submissions.pq` | `Submissions` | yes |
| 8 | `04a_FactInputCell.pq` | `FactInputCell` | yes (hide) |
| 9 | `04b_FactInput.pq` | `FactInput` | yes |
| 10 | `05_FactDeployed.pq` | `FactDeployed` | yes |
| 11 | `06_FactMAPackage.pq` | `FactMAPackage` | yes |
| 12 | `09_FactCalendar.pq` | `FactCalendar` | yes |
| 13 | `10a_FactCommandTeam.pq` | `FactCommandTeam` | yes |
| 14 | `10b_FactOtherResources.pq` | `FactOtherResources` | yes |
| 15 | `08_DataQuality.pq` | `DataQuality` | yes |
| B1–B4 | `11a`–`11d_RouteB_*.pq` | `BackupWorkbooks`, `WB_FactMD`, `WB_FactWT`, `WB_FactOtherPages` | Route B only |

## 3. Tables created in DAX (Modelling → New table)
```
DimDate = CALENDAR ( DATE ( 2026, 9, 1 ), DATE ( 2029, 3, 31 ) )
DimMercurySkill = DISTINCT ( UNION ( VALUES ( DimResource[MercurySkill] ), VALUES ( RefCapacity[MercurySkill] ) ) )
_Measures = { BLANK () }
```

## 4. Relationships (Model view; all many-to-one, single direction)

| Many | One |
|---|---|
| FactDeployed[ResourceID], FactInput[ResourceID], FactMAPackage[ResourceID], FactCalendar[ResourceID] | DimResource[ResourceID] |
| FactDeployed[RankCode], FactCalendar[RankCode], FactCommandTeam[RankCode] | DimRank[RankCode] |
| FactDeployed[RiskCode], FactInput[RiskCode], FactMAPackage[RiskCode], FactCalendar[RiskCode] | DimRisk[RiskCode] |
| FactDeployed[DayCode], FactInput[DayCode], FactMAPackage[DayCode] | DimRelativeDay[DayCode] |
| FactDeployed[ForceKey], FactInput[ForceKey], FactMAPackage[ForceKey], FactCalendar[HostForce], FactCommandTeam[ForceKey], FactOtherResources[ForceKey] | DimForce[Force] |
| FactDeployed[RateKey] | DimRate[RateKey] |
| FactDeployed[SubmissionID], FactInputCell[SubmissionID], DataQuality[SubmissionID] | Submissions[SubmissionID] |
| FactCalendar[Date], FactCommandTeam[Month] | DimDate[Date] |
| RefFixture[MatchDate] | DimDate[Date] (**inactive**) |
| DimResource[MercurySkill], RefCapacity[MercurySkill] | DimMercurySkill[MercurySkill] |
| RefCapacity[Force] / RefCapacity[RankCode] | DimForce / DimRank |
| Route B: WB_FactMD[ResourceID], WB_FactWT[ResourceID] | DimResource[ResourceID] |
| Route B: WB_FactMD[ForceKey], WB_FactWT[ForceKey] | DimForce[Force] |
| Route B: WB_FactMD[Date] | DimDate[Date] |

## 5. Measures
Paste each measure from `measures.dax` into `_Measures` (one per New measure).

## 6. Check it (synthetic sample)
Refresh, then build a matrix with DimRisk[Risk] on columns and compare with
`docs/04_validation.md` §3. All values must match exactly. `DQ Errors` must show 2
and `DQ Warnings` 4 for the synthetic force.

## 7. Report pages
| Page | Main visuals |
|---|---|
| Overview | Grand Total Cost, cards per collection, Forces Submitted, DQ Errors |
| Match Day | matrix Mercury skill × relative day, slicer risk; MD Cost per Fixture |
| Wider Tournament | matrix Mercury skill × risk; WT Tournament Cost |
| Daily deployment | line chart DimDate × Calendar Officers (populates when RefFixture is filled) |
| Budget | cost by portfolio, location (Fan Zone, Town Centre / night-time economy), supply type |
| Capacity | Demand Peak vs Capacity by force / region / skill |
| Equipment / Scenarios | Equipment Required; Officers incl. Scenario |
| Data quality | DataQuality table by force, code, severity |

## Switching rules
`RefParameter[RuleSet]` in the reference workbook: `LEGACY` (default – identical to
the old workbook and the backup workbook) or `AGREED` (revised rules D-02, D-03,
D-05, D-06, D-07). Save, refresh.

## Fixtures
Leave `RefFixture` empty until fixtures are set: Match Day results are then per
fixture by risk level. When fixtures are known, add rows (FixtureID, MatchDate,
Venue, HostForce, RiskCode) and refresh; the daily calendar appears.
