// Query name: Ref_Rates   (do not load)
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    T = Book{[Item = "tblRates", Kind = "Table"]}[Data],
    Typed = Table.TransformColumnTypes(T, {{"Supplied by", type text}, {"Rank", type text}, {"Hourly rate", type number}}),
    Kept = Table.SelectRows(Typed, each [Rank] <> null and [Hourly rate] <> null)
in
    Table.SelectColumns(Kept, {"Supplied by", "Rank", "Hourly rate"})
