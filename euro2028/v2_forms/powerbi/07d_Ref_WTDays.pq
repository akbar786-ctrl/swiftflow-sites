// Query name: Ref_WTDays   (do not load)
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    T = Book{[Item = "tblWTDays", Kind = "Table"]}[Data]
in
    Table.SelectRows(Table.TransformColumnTypes(T, {{"Risk", type text}, {"Days", type number}}), each [Risk] <> null)
