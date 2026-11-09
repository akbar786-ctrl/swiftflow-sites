// Query name: Returns   (do not load)
// Every completed form in ReturnsFolder (Excel temporary "~$" files are ignored).
let
    Files = Folder.Files(ReturnsFolder),
    Xlsx = Table.SelectRows(Files, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),
    Kept = Table.SelectColumns(Xlsx, {"Name", "Content", "Date modified"})
in
    Kept
