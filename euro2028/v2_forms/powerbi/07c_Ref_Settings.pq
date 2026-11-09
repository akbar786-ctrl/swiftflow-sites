// Query name: Ref_Settings   (do not load)
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    T = Book{[Item = "tblSettings", Kind = "Table"]}[Data]
in
    Table.TransformColumnTypes(T, {{"Key", type text}, {"Setting", type text}, {"Value", type number}})
