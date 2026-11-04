// ============================================================================
// fnReadCollection (function)
// Reads ONE collection sheet ("MD Collection" or "WT Collection") of ONE force
// workbook into a long, cell-level table. Relies only on the template contract:
//   * column A  = Row Key  ("<ResourceID>|<Line>" or "SECTION|<coll>-<section>")
//   * hidden key row (column A = "ROWKEY") = Column Key "<Risk>|<Day>|<Measure>"
// so cosmetic changes to the template never break the import.
//
// Output columns:
//   RowKey, ResourceID, Line, RankText, Role, SupplyText, LocationText,
//   RiskCode, DayCode, MeasureCode, RawValue, IsBlank, SectionKey, SectionStatus
// ============================================================================
(content as binary, sheetName as text) as table =>
let
    Book      = Excel.Workbook(content, null, true),
    Sheet     = Book{[Item = sheetName, Kind = "Sheet"]}[Data],
    Indexed   = Table.AddIndexColumn(Sheet, "__Row", 0, 1, Int64.Type),
    KeyRowIdx = Table.SelectRows(Indexed, each [Column1] = "ROWKEY"){0}[__Row],
    KeyRow    = Record.RemoveFields(Indexed{KeyRowIdx}, {"__Row"}),
    OldNames  = Table.ColumnNames(Sheet),
    NewNames  = List.Transform(
                    List.Positions(OldNames),
                    (i) => let k = Record.Field(KeyRow, OldNames{i})
                           in  if k = null or k = "" then "__c" & Text.From(i) else Text.From(k)),
    Renamed   = Table.RenameColumns(Sheet, List.Zip({OldNames, NewNames})),
    Body      = Table.Skip(Renamed, KeyRowIdx + 1),
    WithKey   = Table.SelectRows(Body, each [ROWKEY] <> null and [ROWKEY] <> ""),

    // ---- section status (section rows hold the status in the LOCATION column)
    SectionRows = Table.SelectRows(WithKey, each Text.StartsWith([ROWKEY], "SECTION|")),
    Sections    = Table.TransformColumns(
                      Table.SelectColumns(
                          Table.RenameColumns(SectionRows, {{"ROWKEY", "SectionKey"}, {"LOCATION", "SectionStatus"}}),
                          {"SectionKey", "SectionStatus"}),
                      {{"SectionKey", each Text.AfterDelimiter(_, "|"), type text},
                       {"SectionStatus", each if _ = null then "Not started" else Text.From(_), type text}}),

    // ---- data lines
    Lines       = Table.SelectRows(WithKey, each not Text.StartsWith([ROWKEY], "SECTION|")),
    InputCols   = List.Select(NewNames, each List.Count(Text.Split(_, "|")) = 3 and not Text.StartsWith(_, "TOTAL|")),
    Meta        = {"ROWKEY", "RANK", "ROLE", "SUPPLY", "LOCATION"},
    Kept        = Table.SelectColumns(Lines, Meta & InputCols),
    // Unpivot drops nulls, so mark blanks first to keep blank cells visible
    Marked      = Table.TransformColumns(Kept, List.Transform(InputCols, (c) => {c, each if _ = null or _ = "" then "##BLANK##" else _})),
    Long        = Table.UnpivotOtherColumns(Marked, Meta, "ColKey", "RawValue"),
    Split       = Table.SplitColumn(Long, "ColKey", Splitter.SplitTextByDelimiter("|"), {"RiskCode", "DayCode", "MeasureCode"}),
    Blank       = Table.AddColumn(Split, "IsBlank", each [RawValue] = "##BLANK##", type logical),
    Raw         = Table.TransformColumns(Blank, {{"RawValue", each if _ = "##BLANK##" then null else _}}),
    Keys        = Table.AddColumn(Raw, "ResourceID", each Text.BeforeDelimiter([ROWKEY], "|"), type text),
    LineNo      = Table.AddColumn(Keys, "Line", each Number.From(Text.AfterDelimiter([ROWKEY], "|")), Int64.Type),
    SecKey      = Table.AddColumn(LineNo, "SectionKey", each Text.BeforeDelimiter([ResourceID], "-", 1), type text),
    Renamed2    = Table.RenameColumns(SecKey, {{"ROWKEY", "RowKey"}, {"RANK", "RankText"}, {"ROLE", "Role"},
                                               {"SUPPLY", "SupplyText"}, {"LOCATION", "LocationText"}}),
    WithStatus  = Table.NestedJoin(Renamed2, {"SectionKey"}, Sections, {"SectionKey"}, "S", JoinKind.LeftOuter),
    Result      = Table.ExpandTableColumn(WithStatus, "S", {"SectionStatus"})
in
    Result
