// Query name: fnReadSheet   (function - do not load)
// Reads one force's sheet through its hidden named range (MD_Data or WT_Data, anchored at A1, so Excel row N is
// row N of the grid) and returns one row per Excel row x risk x day with Number, Normal ToD and Extra Hours,
// the text in columns B and C, and the row-map description of that row.
(Content as binary, RangeName as text, RowMap as table, ColMap as table) as table =>
let
    Book = Excel.Workbook(Content, false, true),
    Found = Table.SelectRows(Book, each [Kind] = "DefinedName" and [Name] = RangeName),
    Grid = if Table.RowCount(Found) = 0
        then error Error.Record("Not the EURO 2028 template", "Named range " & RangeName & " not found")
        else Found{0}[Data],
    Indexed = Table.AddIndexColumn(Grid, "Excel row", 1, 1, Int64.Type),
    MappedRows = List.Buffer(RowMap[Excel row]),
    OnlyMapped = Table.SelectRows(Indexed, each List.Contains(MappedRows, [Excel row])),
    NumCols = List.Transform(ColMap[Column number], each "Column" & Text.From(_)),
    Kept = Table.SelectColumns(OnlyMapped, {"Excel row", "Column2", "Column3"} & NumCols, MissingField.UseNull),
    Long = Table.UnpivotOtherColumns(Kept, {"Excel row", "Column2", "Column3"}, "Col", "Raw"),
    Num = Table.AddColumn(Long, "Value", each try Number.From([Raw]) otherwise null, type nullable number),
    NonZero = Table.SelectRows(Num, each [Value] <> null and [Value] <> 0),
    ColNo = Table.AddColumn(NonZero, "Column number", each Number.From(Text.AfterDelimiter([Col], "Column")), Int64.Type),
    WithCol = Table.NestedJoin(ColNo, {"Column number"}, ColMap, {"Column number"}, "CM", JoinKind.Inner),
    Expanded = Table.ExpandTableColumn(WithCol, "CM", {"Risk", "Day", "Measure"}),
    Slim = Table.SelectColumns(Expanded, {"Excel row", "Column2", "Column3", "Risk", "Day", "Measure", "Value"}),
    Wide = Table.Pivot(Slim, {"Number", "Normal ToD", "Extra Hours"}, "Measure", "Value", List.Sum),
    Zeros = Table.TransformColumns(Wide, {
        {"Number", each _ ?? 0, type number}, {"Normal ToD", each _ ?? 0, type number}, {"Extra Hours", each _ ?? 0, type number}}),
    Named = Table.RenameColumns(Zeros, {{"Column2", "Column B"}, {"Column3", "Column C"}}),
    WithMap = Table.NestedJoin(Named, {"Excel row"}, RowMap, {"Excel row"}, "Map", JoinKind.LeftOuter),
    Result = Table.ExpandTableColumn(WithMap, "Map", List.RemoveItems(Table.ColumnNames(RowMap), {"Excel row"}))
in
    Result
