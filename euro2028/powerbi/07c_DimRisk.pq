// ============================================================================
// DimRisk (load: yes) - risk levels plus the national WT day assumption (D-08)
// ============================================================================
let
    Src    = fnRef("DimRisk"),
    Days   = fnRef("RefWTDays"),
    Result = Table.ExpandTableColumn(
                 Table.NestedJoin(Src, {"RiskCode"}, Days, {"RiskCode"}, "D", JoinKind.LeftOuter),
                 "D", {"Days"}, {"WTDays"})
in
    Result
