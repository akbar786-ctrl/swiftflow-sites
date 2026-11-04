// ============================================================================
// FactMAPackage (load: yes) - mutual-aid PSU / serial packages at day grain
// (mirrors engine._deploy, package section). One row per line x risk x day
// with Units entered and PackagesCharged:
//   LEGACY                       per unit per day (legacy Cost Sheet)
//   AGREED  MD matchday-only     PER_UNIT_PER_FIXTURE: units on Match Day
//           MD non-matchday      PER_UNIT_PER_FIXTURE_MAXDAY: units on the peak day only
//           WT                   MAPSU_Basis_WT (default PER_DAY)
// PackageRate from DimMAPackage; missing rate (serials) -> null cost, DQ14.
// ============================================================================
let
    Params   = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    Legacy   = Text.Upper(Text.From(Params[RuleSet])) = "LEGACY",
    MA       = Table.SelectRows(FactInput, each [RowType] = "MA_PSU" and [NUM] <> null and [NUM] > 0),
    WithRes  = Table.ExpandTableColumn(Table.NestedJoin(MA, {"ResourceID"}, DimResource, {"ResourceID"}, "R", JoinKind.LeftOuter),
                   "R", {"MatchdayOnly"}),
    DayIdx   = Table.AddColumn(WithRes, "DayIdx", each try DimRelativeDay{[DayCode = [DayCode]]}[Sort] otherwise 9, Int64.Type),
    Basis    = Table.AddColumn(DayIdx, "Basis", each
                   if Legacy then "PER_UNIT_PER_DAY"
                   else if [Collection] = "WT" then Params[MAPSU_Basis_WT]
                   else if [MatchdayOnly] = true then Params[MAPSU_Basis_Matchday]
                   else Params[MAPSU_Basis_NonMatchday], type text),
    // peak day per line x risk (earliest day wins a tie)
    Peak     = Table.Group(Basis, {"SubmissionID", "RowKey", "RiskCode"},
                   {{"PeakDay", each Table.Sort(_, {{"NUM", Order.Descending}, {"DayIdx", Order.Ascending}}){0}[DayCode], type text}}),
    WithPeak = Table.ExpandTableColumn(Table.NestedJoin(Basis, {"SubmissionID", "RowKey", "RiskCode"}, Peak,
                   {"SubmissionID", "RowKey", "RiskCode"}, "P", JoinKind.LeftOuter), "P", {"PeakDay"}),
    Charged  = Table.AddColumn(WithPeak, "PackagesCharged", each
                   if Text.EndsWith([Basis], "MAXDAY") then (if [DayCode] = [PeakDay] then [NUM] else 0)
                   else if [Basis] = "PER_UNIT_PER_FIXTURE" then (if [DayCode] = "MD" then [NUM] else 0)
                   else [NUM], type number),
    Rate     = Table.ExpandTableColumn(Table.NestedJoin(Charged, {"UnitType"}, DimMAPackage, {"UnitType"}, "K", JoinKind.LeftOuter),
                   "K", {"PackageRate"}),
    Cost     = Table.AddColumn(Rate, "Cost", each if [PackageRate] = null then null else [PackagesCharged] * [PackageRate], type nullable number),
    Result   = Table.SelectColumns(Table.RenameColumns(Cost, {{"NUM", "Units"}}),
                   {"SubmissionID", "ForceKey", "Collection", "ResourceID", "RowKey", "Line", "DayGroup", "RiskCode", "DayCode",
                    "UnitType", "Basis", "Units", "PackagesCharged", "PackageRate", "Cost"}),
    LineKey  = Table.AddColumn(Result, "LineKey", each [RowKey], type text)
in
    LineKey
