// Query name: WT_Lines   (do not load) - every force's Wider Tournament Resource, one row per Excel row x risk
let
    Read = Table.AddColumn(Returns, "Lines", each
        let F = [Force], N = [Name]
        in try Table.Buffer(Table.AddColumn(Table.AddColumn(fnReadSheet([Content], "WT_Data", Ref_RowMapWT, Ref_ColMapWT),
            "Force", (r) => F, type text), "Source file", (r) => N, type text)) otherwise null),
    Ok = List.RemoveNulls(Read[Lines])
in
    if List.IsEmpty(Ok) then error Error.Record("No returns", "No readable EURO 2028 templates in " & ReturnsFolder) else Table.Combine(Ok)
