// ============================================================================
// FactCommandTeam (load: yes) - Route A, from the collection template table
// tblCommandTeam. Legacy method (matches the legacy workbook to the penny):
// monthly cost per post = CTP annual salary / 12 x CTP uplift factor at that
// month; months from Start to End INCLUSIVE.
// ============================================================================
let
    Current  = Table.SelectRows(Submissions, each [IsCurrent]),
    Rows     = Table.Combine(Table.TransformRows(Current, (s) =>
                   let t = try Excel.Workbook(s[Content], true, true){[Item = "tblCommandTeam", Kind = "Table"]}[Data]
                           otherwise #table({"Rank", "Number of posts", "Start month", "End month"}, {})
                   in  Table.AddColumn(Table.AddColumn(t, "SubmissionID", each s[SubmissionID]), "ForceKey", each s[ForceKey]))),
    Valid    = Table.SelectRows(Rows, each [Rank] <> null and [Number of posts] <> null and [Start month] <> null),
    RankCode = Table.AddColumn(Valid, "RankCode", each try DimRank{[Rank = [Rank]]}[RankCode] otherwise null, type text),
    Months   = List.Generate(() => #date(2026, 9, 1), each _ <= #date(2029, 3, 1), each Date.AddMonths(_, 1)),
    Annual   = Table.SelectRows(fnRef("RefAnnualCost"), each [CostTable] = "CTP"),
    Steps    = Table.SelectRows(RefUpliftStep, each [RateTable] = "CTP"),
    Factor   = (m as date) as number => List.Product(List.Combine({{1}, Table.SelectRows(Steps, each Date.From([EffectiveFrom]) <= m)[Factor]})),
    Expanded = Table.ExpandListColumn(Table.AddColumn(RankCode, "Month", each List.Select(Months, (m) =>
                   m >= Date.From([Start month]) and ([End month] = null or m <= Date.From([End month])))), "Month"),
    Cost     = Table.AddColumn(Expanded, "Cost", each
                   let a = try Annual{[RankCode = [RankCode]]}[AnnualCost] otherwise null
                   in  if a = null then null else Number.From([Number of posts]) * a / 12 * Factor([Month]), type nullable number),
    Result   = Table.TransformColumnTypes(Table.SelectColumns(Cost, {"SubmissionID", "ForceKey", "RankCode", "Number of posts", "Month", "Cost"}),
                   {{"Month", type date}})
in
    Result
