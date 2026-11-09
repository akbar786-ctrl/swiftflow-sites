// Query name: Ref_Fixtures   (do not load)
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    T = Book{[Item = "tblFixtures", Kind = "Table"]}[Data],
    Kept = Table.SelectRows(T, each [Date] <> null and [Risk] <> null and [Force] <> null),
    Typed = Table.TransformColumns(Kept, {
        {"Fixture", each if _ = null then "" else Text.From(_), type text},
        {"Date", each Date.From(_), type date},
        {"Risk", each Text.Trim(Text.From(_)), type text},
        {"Force", each Text.Trim(Text.From(_)), type text}})
in
    Typed
