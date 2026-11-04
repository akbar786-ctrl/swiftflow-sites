// ============================================================================
// FactOtherResources (load: yes) - Route A, from the collection template table
// tblOtherResources. Legacy method (matches the legacy workbook to the penny):
// FY 2027-28 hours x HF overtime rate at OtherResourcesRateDate2728 (Jun-2027);
// FY 2028-29 hours x HF overtime rate at RateEffectiveDate (Jun-2028).
// ============================================================================
let
    Params  = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    HF      = (rk as nullable text, d as date) as nullable number =>
                  if rk = null then null else
                  let b = Table.SelectRows(RefBaseRate, each [RateTable] = "HF" and [RankCode] = rk and [RateType] = "OT"),
                      f = List.Product(List.Combine({{1}, Table.SelectRows(RefUpliftStep, each [RateTable] = "HF" and Date.From([EffectiveFrom]) <= d)[Factor]}))
                  in  if Table.IsEmpty(b) then null else b{0}[BaseRate] * f,
    D2728   = Date.From(Params[OtherResourcesRateDate2728]),
    D2829   = Date.From(Params[RateEffectiveDate]),
    Current = Table.SelectRows(Submissions, each [IsCurrent]),
    Rows    = Table.Combine(Table.TransformRows(Current, (s) =>
                  let t = try Excel.Workbook(s[Content], true, true){[Item = "tblOtherResources", Kind = "Table"]}[Data]
                          otherwise #table({"Category", "Rank", "Hours FY 2027-28", "Hours FY 2028-29"}, {})
                  in  Table.AddColumn(Table.AddColumn(t, "SubmissionID", each s[SubmissionID]), "ForceKey", each s[ForceKey]))),
    Valid   = Table.SelectRows(Rows, each [Category] <> null and [Rank] <> null),
    Code    = Table.AddColumn(Valid, "RankCode", each try DimRank{[Rank = [Rank]]}[RateRank] otherwise null, type text),
    Cost    = Table.AddColumn(Table.AddColumn(Code, "Cost FY 2027-28", each
                  let r = HF([RankCode], D2728) in if r = null then null else (if [Hours FY 2027-28] = null then 0 else [Hours FY 2027-28]) * r, type nullable number),
                  "Cost FY 2028-29", each
                  let r = HF([RankCode], D2829) in if r = null then null else (if [Hours FY 2028-29] = null then 0 else [Hours FY 2028-29]) * r, type nullable number)
in
    Cost
