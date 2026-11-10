// Query name: fnRefTable   (function - do not load)
// Returns one named table from the reference workbook.
(TableName as text) as table =>
let
    Book = Excel.Workbook(File.Contents(ReferenceFile), null, true),
    Found = Table.SelectRows(Book, each [Kind] = "Table" and [Item] = TableName)
in
    if Table.RowCount(Found) = 0 then error Error.Record("Reference table missing", TableName) else Found{0}[Data]
