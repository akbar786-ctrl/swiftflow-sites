// Query name: MD_Long   (do not load)
// All forces' MD Collection pages, one row per filled-in cell.
let
    Read = Table.AddColumn(Returns, "Rows", each try fnFormLong([Content], [Name], "tblMatchDay") otherwise null),
    Ok = Table.SelectRows(Read, each [Rows] <> null),
    Combined = if Table.RowCount(Ok) = 0 then null else Table.Combine(Ok[Rows])
in
    Combined
