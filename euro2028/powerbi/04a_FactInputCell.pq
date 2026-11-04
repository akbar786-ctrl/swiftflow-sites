// ============================================================================
// FactInputCell (load: yes, hidden) - every input cell of every CURRENT
// submission, with the blank / zero / N/A / invalid rules applied.
// Mirrors engine._clean_inputs. Feeds FactInput and DataQuality.
// ============================================================================
let
    Params     = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    MaToD      = Text.Upper(Text.From(Params[MA_ToD_Collected])) = "TRUE",
    Current    = Table.SelectRows(Submissions, each [IsCurrent]),
    Read       = Table.AddColumn(Current, "Cells", each
                    Table.Combine({
                        Table.AddColumn(fnReadCollection([Content], "MD Collection"), "Collection", each "MD", type text),
                        Table.AddColumn(fnReadCollection([Content], "WT Collection"), "Collection", each "WT", type text)})),
    Cells      = Table.ExpandTableColumn(Table.SelectColumns(Read, {"SubmissionID", "ForceKey", "Cells"}), "Cells",
                    {"Collection", "RowKey", "ResourceID", "Line", "RankText", "Role", "SupplyText", "LocationText",
                     "RiskCode", "DayCode", "MeasureCode", "RawValue", "IsBlank", "SectionKey", "SectionStatus"}),
    // structural check: unknown row keys are dropped here and reported in DataQuality
    Known      = Table.NestedJoin(Cells, {"RowKey"}, DimResourceLine, {"RowKey"}, "RL", JoinKind.Inner),
    WithLine   = Table.ExpandTableColumn(Known, "RL", {"FixedRankCode", "UnitLevel"}),
    WithRes    = Table.ExpandTableColumn(
                    Table.NestedJoin(WithLine, {"ResourceID"}, DimResource, {"ResourceID"}, "R", JoinKind.LeftOuter),
                    "R", {"RowType", "SupplyCode", "MatchdayOnly", "Portfolio", "Location", "Resource"}),
    LineSupply = Table.AddColumn(WithRes, "LineSupplyCode", each
                    if [RowType] = "ADDITIONAL" then
                        (if [SupplyText] = "Force Supplied" then "FS" else if [SupplyText] = "Mutual Aid Requested" then "MA" else null)
                    else [SupplyCode], type text),
    Applicable = Table.AddColumn(LineSupply, "IsApplicable", each
                    not (([Collection] = "MD" and [MatchdayOnly] = true and [DayCode] <> "MD")
                         or (List.Contains({"PSU_COUNT", "MA_PSU"}, [RowType]) and [MeasureCode] <> "NUM")
                         or ([LineSupplyCode] = "MA" and [MeasureCode] = "TOD" and not MaToD)), type logical),
    Parsed     = Table.AddColumn(Applicable, "Parsed", each
                    if [IsBlank] then null
                    else let n = try Number.From([RawValue]) otherwise null in n, type nullable number),
    State      = Table.AddColumn(Parsed, "ValueState", each
                    if not [IsApplicable] then (if [IsBlank] then "NA" else "NA_WITH_VALUE")
                    else if [IsBlank] then (if List.Contains({"Complete", "Not required"}, [SectionStatus]) then "BLANK_AS_ZERO" else "BLANK")
                    else if [Parsed] = null then "INVALID_TEXT"
                    else if [Parsed] < 0 then "INVALID_NEGATIVE"
                    else if [MeasureCode] = "NUM" and [Parsed] <> Number.Round([Parsed]) then "INVALID_DECIMAL"
                    else if [MeasureCode] <> "NUM" and [Parsed] > 24 then "INVALID_HOURS"
                    else "VALUE", type text),
    Value      = Table.AddColumn(State, "Value", each
                    if [ValueState] = "VALUE" then [Parsed] else if [ValueState] = "BLANK_AS_ZERO" then 0 else null,
                    type nullable number)
in
    Value
