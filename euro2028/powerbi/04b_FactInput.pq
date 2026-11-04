// ============================================================================
// FactInput (load: yes) - one row per line x risk x day with NUM, TOD, XH.
// NOTE: List.Sum over null values returns null (blank stays blank) - required so
// that BLANK and 0 remain distinguishable downstream.
// ============================================================================
let
    Src      = Table.SelectColumns(FactInputCell, {"SubmissionID", "ForceKey", "Collection", "RowKey", "ResourceID",
                  "Line", "RowType", "Portfolio", "Resource", "LineSupplyCode", "FixedRankCode", "UnitLevel",
                  "RankText", "Role", "LocationText", "Location", "SectionKey", "SectionStatus",
                  "RiskCode", "DayCode", "MeasureCode", "Value"}),
    Pivoted  = Table.Pivot(Src, {"NUM", "TOD", "XH"}, "MeasureCode", "Value", List.Sum),
    RankCode = Table.AddColumn(Pivoted, "RankCode", each
                  if [FixedRankCode] <> null then [FixedRankCode]
                  else try DimRank{[Rank = [RankText]]}[RankCode] otherwise null, type text),
    Loc      = Table.AddColumn(RankCode, "LocationFinal", each if [Location] = "Option" then [LocationText] else [Location], type text),
    Typed    = Table.TransformColumnTypes(Loc, {{"NUM", type number}, {"TOD", type number}, {"XH", type number}})
in
    Typed
