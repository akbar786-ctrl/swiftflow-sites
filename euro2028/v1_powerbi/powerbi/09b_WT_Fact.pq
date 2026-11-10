// Query name: WT_Fact   (LOAD) - Wider Tournament: typical day per force / line / risk, and tournament totals
// (typical day x days at that risk level, Ref_WTDays)
let
    Daily = fnBuildFact(WT_Lines),
    J = Table.NestedJoin(Daily, {"Risk"}, Ref_WTDays, {"Risk"}, "D", JoinKind.LeftOuter),
    Days = Table.AddColumn(J, "Days at risk level", each if Table.RowCount([D]) > 0 then [D]{0}[Days] else 0, type number),
    NoD = Table.RemoveColumns(Days, {"D"}),
    T1 = Table.AddColumn(NoD, "Officer days (tournament)", each [Officers] * [Days at risk level], type number),
    T2 = Table.AddColumn(T1, "Overtime cost (tournament)", each [Overtime cost] * [Days at risk level], type number),
    T3 = Table.AddColumn(T2, "Mutual aid PSU cost (tournament)", each [Mutual aid PSU cost] * [Days at risk level], type number),
    T4 = Table.AddColumn(T3, "Normal duty cost (tournament)", each [Normal duty cost] * [Days at risk level], type number),
    T5 = Table.AddColumn(T4, "PSU days (tournament)", each [#"PSUs / Serials"] * [Days at risk level], type number)
in
    T5
