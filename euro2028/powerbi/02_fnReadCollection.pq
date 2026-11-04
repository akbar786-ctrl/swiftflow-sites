// ============================================================================
// fnReadCollection (function) - reads ONE collection sheet of ONE workbook
// ("Match Day Resource" or "Wider Tournament Resource") into a long,
// cell-level table. Relies only on the template contract:
//   * hidden key row: "ROWKEY" over the hidden key column, "COLA"/"COLB"/"COLC"
//     over columns A-C, "<Risk>|<Day>|<Measure>" over input columns;
//   * key column: "<ResourceID>|<Line>" on data rows, "SECTION|..." on section rows.
// Works for both the force collection template and the backup calculation workbook.
//
// Output: RowKey, ResourceID, Line, ColA, ColB, ColC, RiskCode, DayCode,
//         MeasureCode, RawValue, IsBlank
// ============================================================================
(content as binary, sheetName as text) as table =>
let
    Book      = Excel.Workbook(content, null, true),
    Sheet     = Book{[Item = sheetName, Kind = "Sheet"]}[Data],
    Indexed   = Table.AddIndexColumn(Sheet, "__Row", 0, 1, Int64.Type),
    KeyRowIdx = Table.SelectRows(Indexed, each List.Contains(Record.FieldValues(_), "ROWKEY")){0}[__Row],
    KeyRow    = Record.RemoveFields(Indexed{KeyRowIdx}, {"__Row"}),
    OldNames  = Table.ColumnNames(Sheet),
    NewNames  = List.Transform(List.Positions(OldNames), (i) =>
                    let k = Record.Field(KeyRow, OldNames{i})
                    in  if k = null or k = "" then "__c" & Text.From(i) else Text.From(k)),
    Renamed   = Table.RenameColumns(Sheet, List.Zip({OldNames, NewNames})),
    Body      = Table.Skip(Renamed, KeyRowIdx + 1),
    Lines     = Table.SelectRows(Body, each [ROWKEY] <> null and [ROWKEY] <> "" and not Text.StartsWith(Text.From([ROWKEY]), "SECTION|")),
    InputCols = List.Select(NewNames, each List.Count(Text.Split(_, "|")) = 3 and not Text.StartsWith(_, "TOTAL|")),
    Meta      = {"ROWKEY", "COLA", "COLB", "COLC"},
    Kept      = Table.SelectColumns(Lines, Meta & InputCols),
    // Unpivot drops nulls: mark blanks first so blank cells stay visible
    Marked    = Table.TransformColumns(Kept, List.Transform(InputCols, (c) => {c, each if _ = null or _ = "" then "##BLANK##" else _})),
    Long      = Table.UnpivotOtherColumns(Marked, Meta, "ColKey", "RawValue"),
    Split     = Table.SplitColumn(Long, "ColKey", Splitter.SplitTextByDelimiter("|"), {"RiskCode", "DayCode", "MeasureCode"}),
    Blank     = Table.AddColumn(Split, "IsBlank", each [RawValue] = "##BLANK##", type logical),
    Raw       = Table.TransformColumns(Blank, {{"RawValue", each if _ = "##BLANK##" then null else _}}),
    Ids       = Table.AddColumn(Raw, "ResourceID", each Text.BeforeDelimiter([ROWKEY], "|"), type text),
    LineNo    = Table.AddColumn(Ids, "Line", each Number.From(Text.AfterDelimiter([ROWKEY], "|")), Int64.Type),
    Tidy      = Table.TransformColumns(LineNo, {{"COLA", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                                                {"COLB", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                                                {"COLC", each if _ = null then null else Text.Trim(Text.From(_)), type text}}),
    Result    = Table.RenameColumns(Tidy, {{"ROWKEY", "RowKey"}, {"COLA", "ColA"}, {"COLB", "ColB"}, {"COLC", "ColC"}})
in
    Result
