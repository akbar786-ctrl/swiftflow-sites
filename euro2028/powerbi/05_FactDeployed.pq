// ============================================================================
// FactDeployed (load: yes) - officers, hours, rates and cost by
// submission x line x rank x risk x relative day. Mirrors engine._deploy/_cost.
//   Source "Direct"        OFFICER / RANK_SELECT / ROLE / ADDITIONAL rows
//   Source "PSU expansion" force-supplied PSU counts x matching PSU configuration
//                          (same risk and day): officers = units x officers per unit
// Rates (parameter RuleSet):
//   LEGACY  ToD rate 0 (Extra Hours only); Additional Resources at a fixed rate
//           (MD: MA Sergeant, WT: HF Constable); force-supplied In-Stadium PSU
//           Inspectors at the MA Inspector rate; otherwise DimRate overtime rate.
//   AGREED  DimRate ToD and overtime rates for the selected rank.
// Mutual-aid PSUs are costed in FactMAPackage.
// ============================================================================
let
    Params    = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    Legacy    = Text.Upper(Text.From(Params[RuleSet])) = "LEGACY",
    MD1Factor = if Legacy then 1 else Number.From(Params[MD1_PSU_Factor]),
    Z         = (x) => if x = null then 0 else x,
    Officer   = Table.SelectRows(FactInput, each List.Contains({"OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL"}, [RowType]) and Z([NUM]) > 0),
    Direct    = Table.FromRecords(Table.TransformRows(Officer, (r) => [
                    SubmissionID = r[SubmissionID], ForceKey = r[ForceKey], Collection = r[Collection],
                    ResourceID = r[ResourceID], RowKey = r[RowKey], LineKey = r[RowKey], RowType = r[RowType],
                    CostResourceID = r[ResourceID], SupplyCode = r[SupplyCode], RankCode = r[RankCode],
                    Role = r[Role], Description = r[Description], DayGroup = r[DayGroup],
                    RiskCode = r[RiskCode], DayCode = r[DayCode], Source = "Direct", UnitType = null, Units = null,
                    Officers = r[NUM], ToDHours = r[NUM] * Z(r[TOD]), XHHours = r[NUM] * Z(r[XH])])),
    Counts    = Table.SelectRows(FactInput, each [RowType] = "PSU_COUNT" and Z([NUM]) > 0),
    CountsCfg = Table.AddColumn(Counts, "ConfigResourceID", each
                    let res = DimResource{[ResourceID = [ResourceID]]} in if [UnitLevel] = "L2" then res[PsuConfigL2] else res[PsuConfigL3], type text),
    UnitsAdj  = Table.AddColumn(CountsCfg, "UnitsAdj", each [NUM] * (if [DayCode] = "MD+1" then MD1Factor else 1), type number),
    Config    = Table.SelectColumns(Table.SelectRows(FactInput, each [RowType] = "PSU_CONFIG" and Z([NUM]) > 0),
                    {"SubmissionID", "ResourceID", "RiskCode", "DayCode", "RankCode", "NUM", "TOD", "XH"}),
    Joined    = Table.ExpandTableColumn(Table.NestedJoin(UnitsAdj, {"SubmissionID", "ConfigResourceID", "RiskCode", "DayCode"},
                    Config, {"SubmissionID", "ResourceID", "RiskCode", "DayCode"}, "C", JoinKind.Inner),
                    "C", {"RankCode", "NUM", "TOD", "XH"}, {"CfgRank", "CfgNum", "CfgToD", "CfgXH"}),
    Psu       = Table.FromRecords(Table.TransformRows(Joined, (r) => [
                    SubmissionID = r[SubmissionID], ForceKey = r[ForceKey], Collection = r[Collection],
                    ResourceID = r[ResourceID], RowKey = r[RowKey], LineKey = r[RowKey] & "|" & r[CfgRank], RowType = "PSU_COUNT",
                    CostResourceID = r[ConfigResourceID], SupplyCode = "FS", RankCode = r[CfgRank],
                    Role = null, Description = null, DayGroup = r[DayGroup],
                    RiskCode = r[RiskCode], DayCode = r[DayCode], Source = "PSU expansion", UnitType = r[UnitType], Units = r[UnitsAdj],
                    Officers = r[UnitsAdj] * r[CfgNum],
                    ToDHours = r[UnitsAdj] * r[CfgNum] * Z(r[CfgToD]), XHHours = r[UnitsAdj] * r[CfgNum] * Z(r[CfgXH])])),
    Combined  = Table.Combine({Direct, Psu}),
    RateKey   = Table.AddColumn(Combined, "RateKey", each if [RankCode] = null then null else [SupplyCode] & "|" & [RankCode], type text),
    Rates     = Table.AddColumn(RateKey, "R", each
                    let legacyInsp = (try DimResource{[ResourceID = [CostResourceID]]}[LegacyInspMA] otherwise false) = true
                                     and [SupplyCode] = "FS" and (try DimRank{[RankCode = [RankCode]]}[RateRank] otherwise null) = "INSP"
                    in  if Legacy and [RowType] = "ADDITIONAL" then
                            [ToD = 0, OT = if [Collection] = "MD" then Record.Field(DimRate{[RateKey = "MA|SGT"]}, "OTRate")
                                           else Record.Field(DimRate{[RateKey = "FS|PC"]}, "OTRate"), Basis = "Legacy fixed rate"]
                        else if Legacy and legacyInsp then
                            [ToD = 0, OT = Record.Field(DimRate{[RateKey = "MA|INSP"]}, "OTRate"), Basis = "Legacy In-Stadium Inspector"]
                        else if [RateKey] = null then [ToD = null, OT = null, Basis = "No rank"]
                        else let tod = try DimRate{[RateKey = [RateKey]]}[ToDRate] otherwise null,
                                 ot  = try DimRate{[RateKey = [RateKey]]}[OTRate] otherwise null
                             in  [ToD = if Legacy and ot <> null then 0 else tod, OT = ot,
                                  Basis = if ot = null then "No rate" else if Legacy then "Legacy overtime" else "Agreed"]),
    Expanded  = Table.ExpandRecordColumn(Rates, "R", {"ToD", "OT", "Basis"}, {"ToDRate", "OTRate", "RateBasis"}),
    Costs     = Table.AddColumn(Table.AddColumn(Expanded, "ToDCost", each if [ToDRate] = null then null else [ToDHours] * [ToDRate], type nullable number),
                    "XHCost", each if [OTRate] = null then null else [XHHours] * [OTRate], type nullable number),
    Cost      = Table.AddColumn(Costs, "Cost", each if [OTRate] = null then null else Z([ToDCost]) + Z([XHCost]), type nullable number),
    Indexed   = Table.AddIndexColumn(Cost, "DeployedID", 1, 1, Int64.Type),
    Typed     = Table.TransformColumnTypes(Indexed, {{"Units", type number}, {"Officers", type number}, {"ToDHours", type number},
                    {"XHHours", type number}, {"ToDRate", type number}, {"OTRate", type number}})
in
    Typed
