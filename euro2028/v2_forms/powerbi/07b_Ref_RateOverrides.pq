// Query name: Ref_RateOverrides   (do not load)
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    T = Book{[Item = "tblRateOverrides", Kind = "Table"]}[Data],
    Typed = Table.TransformColumnTypes(T, {{"Resource", type text}, {"Supplied by", type text}, {"Rank", type text}, {"Hourly rate", type number}}),
    Kept = Table.SelectRows(Typed, each [Resource] <> null and [Hourly rate] <> null)
in
    Table.RenameColumns(Table.SelectColumns(Kept, {"Resource", "Supplied by", "Rank", "Hourly rate"}), {{"Hourly rate", "Override rate"}})
