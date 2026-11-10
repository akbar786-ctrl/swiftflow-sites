// Query name: MD_Lines   (do not load) - every force's Match Day Resource, one row per Excel row x risk x day
let
    Read = Table.AddColumn(Returns, "Lines", each
        let F = [Force], N = [Name]
        in try Table.Buffer(Table.AddColumn(Table.AddColumn(fnReadSheet([Content], "MD_Data", Ref_RowMapMD, Ref_ColMapMD),
            "Force", (r) => F, type text), "Source file", (r) => N, type text)) otherwise null),
    Ok = List.RemoveNulls(Read[Lines])
in
    if List.IsEmpty(Ok) then error Error.Record("No returns", "No readable EURO 2028 templates in " & ReturnsFolder) else Table.Combine(Ok)
