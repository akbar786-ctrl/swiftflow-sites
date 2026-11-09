// Query name: fnFormLong   (function - do not load)
// Reads one force's form and returns one row per filled-in cell:
//   Force, Source file, Form row, Section, Resource, Supplied by, Role / Skill, Rank, Risk, Day, Measure, Raw, Value
// TableName: "tblMatchDay" (MD Collection page) or "tblWiderTournament" (WT Collection page).
(Content as binary, FileName as text, TableName as text) as table =>
let
    Book = Excel.Workbook(Content, null, true),
    DetailsRows = Table.SelectRows(Book, each [Kind] = "Table" and [Item] = "tblDetails"),
    Details = if Table.RowCount(DetailsRows) = 0 then null else DetailsRows{0}[Data],
    TypedForce = if Details = null or Table.RowCount(Details) = 0 then null else Record.FieldOrDefault(Details{0}, "Force", null),
    ForceText = if TypedForce = null then "" else Text.Trim(Text.From(TypedForce)),
    // if the force name was left blank, use the file name
    Force = if ForceText = "" then Text.BeforeDelimiter(FileName, ".", {0, RelativePosition.FromEnd}) else ForceText,
    DataRows = Table.SelectRows(Book, each [Kind] = "Table" and [Item] = TableName),
    Data = if Table.RowCount(DataRows) = 0
        then error Error.Record("Not a collection form", "Table " & TableName & " was not found in " & FileName)
        else DataRows{0}[Data],
    KeyCols = {"Section", "Resource", "Supplied by", "Role / Skill", "Rank"},
    ValueCols = List.Select(Table.ColumnNames(Data), each Text.Contains(_, " | ")),
    Slim = Table.SelectColumns(Data, KeyCols & ValueCols),
    // "Form row" numbers each line of the form (1 = first line under the headings), so two lines with the
    // same resource / role / rank are never merged
    Indexed = Table.AddIndexColumn(Slim, "Form row", 1, 1, Int64.Type),
    // one row per filled-in cell (blank cells are dropped)
    Long = Table.UnpivotOtherColumns(Indexed, KeyCols & {"Form row"}, "Attribute", "Raw"),
    Parts = Table.AddColumn(Long, "Parts", each Text.Split([Attribute], " | ")),
    Risk = Table.AddColumn(Parts, "Risk", each [Parts]{0}, type text),
    Day = Table.AddColumn(Risk, "Day", each if List.Count([Parts]) = 3 then [Parts]{1} else "WT", type text),
    Measure = Table.AddColumn(Day, "Measure", each List.Last([Parts]), type text),
    Value = Table.AddColumn(Measure, "Value", each try Number.From([Raw]) otherwise null, type nullable number),
    RawText = Table.TransformColumns(Value, {{"Raw", each if _ = null then null else Text.From(_), type nullable text}}),
    WithForce = Table.AddColumn(RawText, "Force", each Force, type text),
    WithFile = Table.AddColumn(WithForce, "Source file", each FileName, type text),
    Result = Table.SelectColumns(WithFile,
        {"Force", "Source file", "Form row"} & KeyCols & {"Risk", "Day", "Measure", "Raw", "Value"}),
    Typed = Table.TransformColumnTypes(Result, {
        {"Section", type text}, {"Resource", type text}, {"Supplied by", type text},
        {"Role / Skill", type text}, {"Rank", type text}})
in
    Typed
