// ============================================================================
// WB_FactMD (load: yes) - Route B Match Day officers and cost by line and
// date, read from the backup workbook pages "MD Demand Template" and
// "MD Cost Template" (hidden row 3 holds ISO dates, column A the line key).
// Equivalent to FactCalendar summed by LineKey and Date under RuleSet LEGACY.
// ============================================================================
let
    Read     = (content as binary, sheet as text, valueName as text) as table =>
                   let ws    = Excel.Workbook(content, false, true){[Item = sheet, Kind = "Sheet"]}[Data],
                       idx   = Table.AddIndexColumn(ws, "__r", 0, 1),
                       keyR  = Table.SelectRows(idx, each [Column1] = "DATEKEY"){0}[__r],
                       names = List.Transform(Record.FieldValues(Record.RemoveFields(idx{keyR}, {"__r"})),
                                   each if _ = null then null else Text.From(_)),
                       cols  = Table.ColumnNames(ws),
                       dated = List.Select(List.Positions(cols), (i) => names{i} <> null and names{i} <> "DATEKEY"),
                       body  = Table.SelectRows(Table.Skip(ws, keyR + 2), each [Column1] <> null),
                       kept  = Table.SelectColumns(body, {"Column1"} & List.Transform(dated, each cols{_})),
                       ren   = Table.RenameColumns(kept, {{"Column1", "LineKey"}} & List.Transform(dated, each {cols{_}, names{_}})),
                       long  = Table.UnpivotOtherColumns(ren, {"LineKey"}, "Date", valueName)
                   in  Table.TransformColumnTypes(long, {{"Date", type date}, {valueName, type number}}),
    Current  = Table.SelectRows(BackupWorkbooks, each [IsCurrent]),
    Rows     = Table.Combine(Table.TransformRows(Current, (w) =>
                   let d = Read(w[Content], "MD Demand Template", "Officers"),
                       c = Read(w[Content], "MD Cost Template", "Cost"),
                       j = Table.ExpandTableColumn(Table.NestedJoin(d, {"LineKey", "Date"}, c, {"LineKey", "Date"}, "C", JoinKind.FullOuter), "C", {"Cost"})
                   in  Table.AddColumn(j, "ForceKey", each w[ForceKey]))),
    ResId    = Table.AddColumn(Rows, "ResourceID", each Text.BeforeDelimiter([LineKey], "|"), type text),
    Rank     = Table.AddColumn(ResId, "RankCode", each
                   let parts = Text.Split([LineKey], "|") in if List.Count(parts) = 3 then parts{2}
                   else try DimResourceLine{[RowKey = [LineKey]]}[FixedRankCode] otherwise null, type text),
    NonZero  = Table.SelectRows(Rank, each ([Officers] <> null and [Officers] <> 0) or ([Cost] <> null and [Cost] <> 0))
in
    NonZero
