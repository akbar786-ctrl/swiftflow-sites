// ============================================================================
// FactDeployed (load: yes) - officers and hours by force x resource x rank x
// risk x relative day. Two sources (mirrors engine._deploy):
//   Direct        - OFFICER / RANK_SELECT / ROLE / ADDITIONAL rows
//   PSU expansion - force-supplied PSU / serial counts x matching PSU
//                   configuration (officers per unit, hours per officer)
// Mutual-aid PSUs are NOT here: they are costed by package in FactMAPackage.
// RateKey = SupplyCode|RankCode joins to DimRate. Costs are DAX measures.
// ============================================================================
let
    Params      = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    MD1Factor   = Number.From(Params[MD1_PSU_Factor]),
    Officer     = Table.SelectRows(FactInput, each List.Contains({"OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL"}, [RowType])
                                              and [NUM] <> null and [NUM] > 0),
    Direct      = Table.FromRecords(Table.TransformRows(Officer, (r) => [
                      SubmissionID = r[SubmissionID], ForceKey = r[ForceKey], Collection = r[Collection],
                      ResourceID = r[ResourceID], RowKey = r[RowKey], SupplyCode = r[LineSupplyCode],
                      RankCode = r[RankCode], Role = r[Role], Location = r[LocationFinal],
                      RiskCode = r[RiskCode], DayCode = r[DayCode], Source = "Direct", UnitType = null, Units = null,
                      Officers = r[NUM],
                      ToDHours = r[NUM] * (if r[TOD] = null then 0 else r[TOD]),
                      XHHours  = r[NUM] * (if r[XH]  = null then 0 else r[XH])])),

    // ---- PSU expansion
    Counts      = Table.SelectRows(FactInput, each [RowType] = "PSU_COUNT" and [NUM] <> null and [NUM] > 0),
    CountsCfg   = Table.AddColumn(Counts, "ConfigResourceID", each
                      let res = DimResource{[ResourceID = [ResourceID]]}
                      in  if [UnitLevel] = "L2" then res[PsuConfigL2] else res[PsuConfigL3], type text),
    Units       = Table.AddColumn(CountsCfg, "UnitsAdj", each [NUM] * (if [DayCode] = "MD+1" then MD1Factor else 1), type number),
    Config      = Table.SelectRows(FactInput, each [RowType] = "PSU_CONFIG" and [NUM] <> null and [NUM] > 0),
    CfgSlim     = Table.SelectColumns(Config, {"SubmissionID", "ResourceID", "RiskCode", "DayCode", "RankCode", "NUM", "TOD", "XH"}),
    Joined      = Table.NestedJoin(Units, {"SubmissionID", "ConfigResourceID", "RiskCode", "DayCode"},
                                   CfgSlim, {"SubmissionID", "ResourceID", "RiskCode", "DayCode"}, "Cfg", JoinKind.Inner),
    Expanded    = Table.ExpandTableColumn(Joined, "Cfg", {"RankCode", "NUM", "TOD", "XH"},
                                          {"CfgRank", "CfgNum", "CfgToD", "CfgXH"}),
    Psu         = Table.FromRecords(Table.TransformRows(Expanded, (r) => [
                      SubmissionID = r[SubmissionID], ForceKey = r[ForceKey], Collection = r[Collection],
                      ResourceID = r[ResourceID], RowKey = r[RowKey], SupplyCode = "FS",
                      RankCode = r[CfgRank], Role = null, Location = r[LocationFinal],
                      RiskCode = r[RiskCode], DayCode = r[DayCode], Source = "PSU expansion",
                      UnitType = r[RankText], Units = r[UnitsAdj],
                      Officers = r[UnitsAdj] * r[CfgNum],
                      ToDHours = r[UnitsAdj] * r[CfgNum] * (if r[CfgToD] = null then 0 else r[CfgToD]),
                      XHHours  = r[UnitsAdj] * r[CfgNum] * (if r[CfgXH]  = null then 0 else r[CfgXH])])),
    Combined    = Table.Combine({Direct, Psu}),
    RateKey     = Table.AddColumn(Combined, "RateKey", each
                      if [SupplyCode] = null or [RankCode] = null then null else [SupplyCode] & "|" & [RankCode], type text),
    DeployKey   = Table.AddIndexColumn(RateKey, "DeployedID", 1, 1, Int64.Type),
    Typed       = Table.TransformColumnTypes(DeployKey, {{"Units", type number}, {"Officers", type number},
                      {"ToDHours", type number}, {"XHHours", type number}, {"SupplyCode", type text}, {"RankCode", type text},
                      {"RiskCode", type text}, {"DayCode", type text}, {"Location", type text}, {"UnitType", type text}})
in
    Typed
