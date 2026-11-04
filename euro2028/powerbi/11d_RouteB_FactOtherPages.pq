// ============================================================================
// WB_FactOtherPages (load: yes) - Route B Command Team & Planning (monthly
// cost by rank) and Other Resources (cost by category, rank and year) read
// from the backup workbook, in one long table:
//   Page, ForceKey, Category, Rank, Period, Cost
// ============================================================================
let
    Current = Table.SelectRows(BackupWorkbooks, each [IsCurrent]),
    Ctp     = (w) =>
                 let ws   = Excel.Workbook(w[Content], false, true){[Item = "Command Team & Planning", Kind = "Sheet"]}[Data],
                     idx  = Table.AddIndexColumn(ws, "__r", 0, 1),
                     hr   = Table.SelectRows(idx, each [Column2] = "Rank" and [Column3] = "Posts"){0}[__r],
                     t    = Table.PromoteHeaders(Table.Skip(ws, hr), [PromoteAllScalars = true]),
                     body = Table.SelectRows(Table.RemoveColumns(t, {"Column1"}, MissingField.Ignore),
                                each [Rank] <> null and [Rank] <> "TOTAL"),
                     mcols = List.Select(Table.ColumnNames(body), each not List.Contains({"Rank", "Posts", "Start", "End", "TOTAL"}, _)),
                     long = Table.UnpivotOtherColumns(Table.SelectColumns(body, {"Rank"} & mcols), {"Rank"}, "Period", "Cost")
                 in  Table.AddColumn(Table.AddColumn(Table.AddColumn(long, "Page", each "Command Team & Planning"),
                         "ForceKey", each w[ForceKey]), "Category", each "Command Team & Planning"),
    Oth     = (w) =>
                 let ws   = Excel.Workbook(w[Content], false, true){[Item = "Other Resources", Kind = "Sheet"]}[Data],
                     idx  = Table.AddIndexColumn(ws, "__r", 0, 1),
                     hr   = Table.SelectRows(idx, each [Column1] = "Category"){0}[__r],
                     t    = Table.SelectRows(Table.PromoteHeaders(Table.Skip(ws, hr), [PromoteAllScalars = true]),
                                each [Category] <> null and [Category] <> "TOTAL"),
                     long = Table.UnpivotOtherColumns(Table.SelectColumns(t, {"Category", "Rank", "Cost FY 2027-28", "Cost FY 2028-29"}),
                                {"Category", "Rank"}, "Period", "Cost")
                 in  Table.AddColumn(Table.AddColumn(long, "Page", each "Other Resources"), "ForceKey", each w[ForceKey]),
    All     = Table.Combine(Table.TransformRows(Current, each Table.Combine({Ctp(_), Oth(_)}))),
    Typed   = Table.TransformColumnTypes(All, {{"Cost", type number}, {"Period", type text}})
in
    Typed
