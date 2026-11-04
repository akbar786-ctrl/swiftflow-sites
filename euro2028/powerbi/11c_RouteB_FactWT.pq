// ============================================================================
// WB_FactWT (load: yes) - Route B Wider Tournament officers per day, cost per
// day, national days and tournament cost by line and risk level, read from
// "WT Demand Template" and "WT Cost Template" of the backup workbook.
// ============================================================================
let
    Risks    = {{"Low", "L"}, {"Medium", "M"}, {"High", "H"}, {"Super High", "SH"}},
    ReadPage = (content as binary, sheet as text) as table =>
                   let ws  = Excel.Workbook(content, false, true){[Item = sheet, Kind = "Sheet"]}[Data],
                       idx = Table.AddIndexColumn(ws, "__r", 0, 1),
                       hr  = Table.SelectRows(idx, each [Column1] = "Row key"){0}[__r],
                       t   = Table.PromoteHeaders(Table.Skip(ws, hr), [PromoteAllScalars = true])
                   in  Table.SelectRows(t, each [Row key] <> null),
    Current  = Table.SelectRows(BackupWorkbooks, each [IsCurrent]),
    Rows     = Table.Combine(Table.TransformRows(Current, (w) =>
                   let d = ReadPage(w[Content], "WT Demand Template"),
                       c = ReadPage(w[Content], "WT Cost Template"),
                       perRisk = List.Transform(Risks, (rk) =>
                           Table.FromRecords(List.Transform(Table.ToRecords(d), (r) =>
                               let cr = Table.SelectRows(c, each [Row key] = r[Row key]){0}
                               in  [ForceKey = w[ForceKey], LineKey = r[Row key], RiskCode = rk{1},
                                    OfficersPerDay = Record.Field(r, rk{0}),
                                    CostPerDay = Record.Field(cr, "COST PER DAY - " & Text.Upper(rk{0})),
                                    Days = Record.Field(cr, "DAYS - " & Text.Upper(rk{0}))])))
                   in  Table.Combine(perRisk))),
    Typed    = Table.TransformColumnTypes(Rows, {{"OfficersPerDay", type number}, {"CostPerDay", type number}, {"Days", type number}}),
    Tourn    = Table.AddColumn(Typed, "TournamentCost", each (if [CostPerDay] = null then 0 else [CostPerDay]) * (if [Days] = null then 0 else [Days]), type number),
    ResId    = Table.AddColumn(Tourn, "ResourceID", each Text.BeforeDelimiter([LineKey], "|"), type text)
in
    ResId
