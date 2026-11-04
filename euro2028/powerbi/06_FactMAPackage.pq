// ============================================================================
// FactMAPackage (load: yes) - mutual-aid PSU / serial packages
// (mirrors engine._deploy, package section). Package basis per parameter:
//   MAPSU_Basis_Matchday     PER_UNIT_PER_FIXTURE        units on MD
//   MAPSU_Basis_NonMatchday  PER_UNIT_PER_FIXTURE_MAXDAY peak units across MD-2..MD+1
//                            PER_UNIT_PER_DAY            sum of units across days (legacy)
//   MAPSU_Basis_WT           PER_DAY                     units per typical WT day
// PackageRate comes from DimMAPackage (07_DimRate.pq). Missing rate -> null
// cost and a DQ14 row in DataQuality.
// ============================================================================
let
    Params    = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    MA        = Table.SelectRows(FactInput, each [RowType] = "MA_PSU" and [NUM] <> null and [NUM] > 0),
    WithRes   = Table.ExpandTableColumn(Table.NestedJoin(MA, {"ResourceID"}, DimResource, {"ResourceID"}, "R", JoinKind.LeftOuter),
                                        "R", {"MatchdayOnly"}),
    Grouped   = Table.Group(WithRes, {"SubmissionID", "ForceKey", "Collection", "ResourceID", "Line", "RiskCode", "MatchdayOnly"},
                    {{"UnitType", each List.First([RankText]), type text},
                     {"Rows", each _, type table}}),
    Basis     = Table.AddColumn(Grouped, "Basis", each
                    if [Collection] = "WT" then Params[MAPSU_Basis_WT]
                    else if [MatchdayOnly] = true then Params[MAPSU_Basis_Matchday]
                    else Params[MAPSU_Basis_NonMatchday], type text),
    Packages  = Table.AddColumn(Basis, "Packages", each
                    if Text.EndsWith([Basis], "MAXDAY") then List.Max([Rows][NUM])
                    else if [Basis] = "PER_UNIT_PER_FIXTURE" then List.Sum(Table.SelectRows([Rows], each [DayCode] = "MD")[NUM])
                    else List.Sum([Rows][NUM]), type number),
    Rate      = Table.ExpandTableColumn(Table.NestedJoin(Packages, {"UnitType"}, DimMAPackage, {"UnitType"}, "P", JoinKind.LeftOuter),
                                        "P", {"PackageRate"}),
    Cost      = Table.AddColumn(Rate, "PackageCost", each if [PackageRate] = null then null else [Packages] * [PackageRate], type nullable number),
    Days      = Table.AddColumn(Cost, "UnitsByDay", each Text.Combine(List.Transform(Table.ToRecords([Rows]),
                    (r) => r[DayCode] & "=" & Text.From(r[NUM])), "; "), type text),
    // Anchor the package to a relative day so day slicers do not repeat it:
    // MD packages sit on Match Day (the package covers travel in, MD, travel out);
    // WT packages sit on the typical WT day.
    WithDay   = Table.AddColumn(Days, "DayCode", each if [Collection] = "WT" then "WT" else "MD", type text),
    Result    = Table.RemoveColumns(WithDay, {"Rows"})
in
    Result
