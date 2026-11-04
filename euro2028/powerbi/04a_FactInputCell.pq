// ============================================================================
// FactInputCell (load: yes, hidden) - every input cell of every CURRENT
// submission with the blank / zero / N-A / invalid rules applied.
// Mirrors engine._clean_inputs. Feeds FactInput and DataQuality.
//
// N/A rules (RefNARule): NA-01 matchday-only -> Match Day only; NA-04
// non-matchday -> Match Day N/A; NA-02 PSU count rows -> Number only;
// NA-03 mutual-aid Normal ToD N/A unless parameter MA_ToD_Collected = TRUE.
// ============================================================================
let
    Params     = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    MaToD      = Text.Upper(Text.From(Params[MA_ToD_Collected])) = "TRUE",
    Current    = Table.SelectRows(Submissions, each [IsCurrent]),
    // section completion (Cover sheet table tblSectionStatus)
    Status     = Table.Combine(Table.TransformRows(Current, (s) =>
                    let t = try Excel.Workbook(s[Content], true, true){[Item = "tblSectionStatus", Kind = "Table"]}[Data]
                            otherwise #table({"SectionKey", "Status"}, {})
                    in  Table.AddColumn(Table.SelectColumns(t, {"SectionKey", "Status"}), "SubmissionID", each s[SubmissionID]))),
    Read       = Table.AddColumn(Current, "Cells", each Table.Combine({
                    Table.AddColumn(fnReadCollection([Content], "Match Day Resource"), "Collection", each "MD", type text),
                    Table.AddColumn(fnReadCollection([Content], "Wider Tournament Resource"), "Collection", each "WT", type text)})),
    Cells      = Table.ExpandTableColumn(Table.SelectColumns(Read, {"SubmissionID", "ForceKey", "Cells"}), "Cells",
                    {"Collection", "RowKey", "ResourceID", "Line", "ColA", "ColB", "ColC", "RiskCode", "DayCode",
                     "MeasureCode", "RawValue", "IsBlank"}),
    Known      = Table.ExpandTableColumn(Table.NestedJoin(Cells, {"RowKey"}, DimResourceLine, {"RowKey"}, "RL", JoinKind.Inner),
                    "RL", {"FixedRankCode", "UnitLevel"}),
    WithRes    = Table.ExpandTableColumn(Table.NestedJoin(Known, {"ResourceID"}, DimResource, {"ResourceID"}, "R", JoinKind.LeftOuter),
                    "R", {"RowType", "SupplyCode", "MatchdayOnly", "DayGroup", "SectionCode"}),
    SecKey     = Table.AddColumn(WithRes, "SectionKey", each [Collection] & "-" & [SectionCode], type text),
    WithStatus = Table.ExpandTableColumn(Table.NestedJoin(SecKey, {"SubmissionID", "SectionKey"}, Status,
                    {"SubmissionID", "SectionKey"}, "S", JoinKind.LeftOuter), "S", {"Status"}, {"SectionStatus"}),
    StatusFix  = Table.TransformColumns(WithStatus, {{"SectionStatus", each if _ = null then "Not started" else _, type text}}),
    Applicable = Table.AddColumn(StatusFix, "IsApplicable", each
                    not (([Collection] = "MD" and [MatchdayOnly] = true and [DayCode] <> "MD")
                         or ([Collection] = "MD" and [DayGroup] = "NONMD" and [DayCode] = "MD")
                         or (List.Contains({"PSU_COUNT", "MA_PSU"}, [RowType]) and [MeasureCode] <> "NUM")
                         or ([SupplyCode] = "MA" and [MeasureCode] = "TOD" and not MaToD)), type logical),
    Parsed     = Table.AddColumn(Applicable, "Parsed", each
                    if [IsBlank] then null else try Number.From([RawValue]) otherwise null, type nullable number),
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
