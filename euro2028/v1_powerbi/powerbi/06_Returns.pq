// Query name: Returns   (do not load)
// Every .xlsx in ReturnsFolder; Force = file name without ".xlsx".
let
    Files = Folder.Files(ReturnsFolder),
    Xlsx = Table.SelectRows(Files, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),
    WithForce = Table.AddColumn(Xlsx, "Force", each Text.Trim(Text.BeforeDelimiter([Name], ".", {0, RelativePosition.FromEnd})), type text)
in
    Table.SelectColumns(WithForce, {"Force", "Name", "Content"})
