// Query name: MD_Fact   (LOAD)
// One row per Force / line on the form / Risk / Day, with officers, hours and costs.
// Costing follows the original collection workbook:
//   overtime cost   = Officers x Extra hrs x hourly rate (rank + supplied by, or a resource override)
//   normal duty cost = Officers x Normal hrs x hourly rate (shown separately)
//   mutual aid PSU / Serial = number of units x package per unit per day
let
    Filled = Table.SelectRows(MD_Long, each [Value] <> null and [Value] <> 0),
    Slim = Table.RemoveColumns(Filled, {"Raw"}),
    Wide = Table.Pivot(Slim, {"Officers", "Normal hrs", "Extra hrs"}, "Measure", "Value", List.Sum),
    Units = Table.AddColumn(Wide, "Is units line", each [Rank] = "Number of PSUs / Serials", type logical),
    PSUs = Table.AddColumn(Units, "PSUs / Serials", each if [Is units line] then ([Officers] ?? 0) else 0, type number),
    Officers = Table.AddColumn(PSUs, "Officer count", each if [Is units line] then 0 else ([Officers] ?? 0), type number),
    Hours = Table.TransformColumns(Officers, {{"Normal hrs", each _ ?? 0, type number}, {"Extra hrs", each _ ?? 0, type number}}),
    Renamed = Table.RenameColumns(Table.RemoveColumns(Hours, {"Officers"}), {{"Officer count", "Officers"}}),
    // hourly rate: resource override first, otherwise rank + supplied by
    WithOverride = Table.NestedJoin(Renamed, {"Resource", "Supplied by", "Rank"}, Ref_RateOverrides, {"Resource", "Supplied by", "Rank"}, "Ovr", JoinKind.LeftOuter),
    WithBase = Table.NestedJoin(WithOverride, {"Supplied by", "Rank"}, Ref_Rates, {"Supplied by", "Rank"}, "Base", JoinKind.LeftOuter),
    Rate = Table.AddColumn(WithBase, "Hourly rate", each
        if Table.RowCount([Ovr]) > 0 then [Ovr]{0}[Override rate]
        else if Table.RowCount([Base]) > 0 then [Base]{0}[Hourly rate]
        else null, type nullable number),
    NoJoins = Table.RemoveColumns(Rate, {"Ovr", "Base"}),
    Package = Ref_Settings{[Key = "MA_PSU_PACKAGE"]}[Value],
    ExtraHours = Table.AddColumn(NoJoins, "Extra hours total", each [Officers] * [Extra hrs], type number),
    NormalHours = Table.AddColumn(ExtraHours, "Normal hours total", each [Officers] * [Normal hrs], type number),
    Overtime = Table.AddColumn(NormalHours, "Overtime cost", each [Extra hours total] * ([Hourly rate] ?? 0), type number),
    NormalCost = Table.AddColumn(Overtime, "Normal duty cost", each [Normal hours total] * ([Hourly rate] ?? 0), type number),
    PSUCost = Table.AddColumn(NormalCost, "Mutual aid PSU cost", each
        if [Is units line] and [Supplied by] = "Mutual aid" then [#"PSUs / Serials"] * Package else 0, type number),
    // lines for days other than Match Day follow a different day rule on the fixture calendar (original workbook)
    RowType = Table.AddColumn(PSUCost, "Row type", each if Text.Contains([Resource] ?? "", "(Non-Matchday)") then "NON_MD" else "STD", type text),
    Typed = Table.TransformColumnTypes(RowType, {{"Risk", type text}, {"Day", type text}, {"Force", type text}})
in
    Typed
