// Query name: WT_Fact   (LOAD)
// One row per Force / line on the form / Risk for a typical Wider Tournament day,
// plus tournament totals = daily figure x number of days at that risk level (Ref_WTDays).
let
    Filled = Table.SelectRows(WT_Long, each [Value] <> null and [Value] <> 0),
    Slim = Table.RemoveColumns(Filled, {"Raw", "Day"}),
    Wide = Table.Pivot(Slim, {"Officers", "Normal hrs", "Extra hrs"}, "Measure", "Value", List.Sum),
    Units = Table.AddColumn(Wide, "Is units line", each [Rank] = "Number of PSUs / Serials", type logical),
    PSUs = Table.AddColumn(Units, "PSUs / Serials", each if [Is units line] then ([Officers] ?? 0) else 0, type number),
    Officers = Table.AddColumn(PSUs, "Officer count", each if [Is units line] then 0 else ([Officers] ?? 0), type number),
    Hours = Table.TransformColumns(Officers, {{"Normal hrs", each _ ?? 0, type number}, {"Extra hrs", each _ ?? 0, type number}}),
    Renamed = Table.RenameColumns(Table.RemoveColumns(Hours, {"Officers"}), {{"Officer count", "Officers"}}),
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
    WithDays = Table.NestedJoin(PSUCost, {"Risk"}, Ref_WTDays, {"Risk"}, "WTDays", JoinKind.LeftOuter),
    Days = Table.AddColumn(WithDays, "Days at risk level", each if Table.RowCount([WTDays]) > 0 then [WTDays]{0}[Days] else 0, type number),
    NoDays = Table.RemoveColumns(Days, {"WTDays"}),
    TOfficers = Table.AddColumn(NoDays, "Officer days (tournament)", each [Officers] * [Days at risk level], type number),
    TOvertime = Table.AddColumn(TOfficers, "Overtime cost (tournament)", each [Overtime cost] * [Days at risk level], type number),
    TNormal = Table.AddColumn(TOvertime, "Normal duty cost (tournament)", each [Normal duty cost] * [Days at risk level], type number),
    TPSU = Table.AddColumn(TNormal, "Mutual aid PSU cost (tournament)", each [Mutual aid PSU cost] * [Days at risk level], type number),
    Typed = Table.TransformColumnTypes(TPSU, {{"Risk", type text}, {"Force", type text}})
in
    Typed
