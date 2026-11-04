// ============================================================================
// DimRate (load: yes) - effective hourly rates at RateEffectiveDate, one row
// per SupplyCode x RankCode (RateKey = "FS|SGT"). Mirrors engine.Rates:
//   rate = BaseRate (RateTable HF for FS, MA for MA, by RateRank)
//          x product of RefUpliftStep[Factor] where EffectiveFrom <= RateEffectiveDate
//   RefZeroCost overrides to 0 (e.g. force-supplied Inspector and above).
//   No base rate -> null (reported as DQ14).
// ============================================================================
let
    Params    = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    EffDate   = Date.From(Params[RateEffectiveDate]),
    Factor    = (tbl as text) as number =>
                    List.Product(List.Combine({{1},
                        Table.SelectRows(RefUpliftStep, each [RateTable] = tbl and Date.From([EffectiveFrom]) <= EffDate)[Factor]})),
    Supplies  = Table.FromRecords({[SupplyCode = "FS", RateTable = "HF"], [SupplyCode = "MA", RateTable = "MA"]}),
    Cross     = Table.ExpandTableColumn(Table.AddColumn(Supplies, "R", each DimRank), "R", {"RankCode", "Rank", "RateRank"}),
    Rate      = (tbl as text, rk as text, typ as text) as nullable number =>
                    let rows = Table.SelectRows(RefBaseRate, each [RateTable] = tbl and [RankCode] = rk and [RateType] = typ)
                    in  if Table.IsEmpty(rows) then null else rows{0}[BaseRate] * Factor(tbl),
    IsZero    = (s as text, rk as text) as logical =>
                    not Table.IsEmpty(Table.SelectRows(RefZeroCost, each [SupplyCode] = s and [RankCode] = rk)),
    WithRates = Table.AddColumn(Cross, "Rates", each
                    if IsZero([SupplyCode], [RateRank]) then [ToDRate = 0, OTRate = 0, RateNote = "GBP 0 rule"]
                    else [ToDRate = Rate([RateTable], [RateRank], "TOD"), OTRate = Rate([RateTable], [RateRank], "OT"),
                          RateNote = if Rate([RateTable], [RateRank], "OT") = null then "NO RATE" else [RateTable] & " table"]),
    Expanded  = Table.ExpandRecordColumn(WithRates, "Rates", {"ToDRate", "OTRate", "RateNote"}),
    Key       = Table.AddColumn(Expanded, "RateKey", each [SupplyCode] & "|" & [RankCode], type text),
    Typed     = Table.TransformColumnTypes(Key, {{"ToDRate", type number}, {"OTRate", type number}})
in
    Typed
