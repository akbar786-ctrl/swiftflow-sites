// ============================================================================
// FactInput (load: yes) - one row per submission x line x risk x day, with
// NUM, TOD, XH and the line attributes chosen by the force:
//   RankCode  fixed rank, or the rank chosen (column A on role rows, C otherwise)
//   Role      role profile (column C on role rows)
//   Description  free text (column B on Additional Resource rows)
//   UnitType  PSU / serial choice (column C on PSU count rows)
// List.Sum over nulls returns null, so BLANK and 0 stay distinguishable.
// ============================================================================
let
    Src      = Table.SelectColumns(FactInputCell, {"SubmissionID", "ForceKey", "Collection", "RowKey", "ResourceID", "Line",
                  "RowType", "SupplyCode", "DayGroup", "FixedRankCode", "UnitLevel", "ColA", "ColB", "ColC",
                  "SectionKey", "SectionStatus", "RiskCode", "DayCode", "MeasureCode", "Value"}),
    Pivoted  = Table.Pivot(Src, {"NUM", "TOD", "XH"}, "MeasureCode", "Value", List.Sum),
    RankText = Table.AddColumn(Pivoted, "RankText", each
                  if [FixedRankCode] <> null then null else if [RowType] = "ROLE" then [ColA]
                  else if List.Contains({"RANK_SELECT", "ADDITIONAL"}, [RowType]) then [ColC] else null, type text),
    RankCode = Table.AddColumn(RankText, "RankCode", each
                  if [FixedRankCode] <> null then [FixedRankCode]
                  else if [RankText] = null then null
                  else try DimRank{[Rank = [RankText]]}[RankCode] otherwise null, type text),
    Role     = Table.AddColumn(RankCode, "Role", each if [RowType] = "ROLE" then [ColC] else null, type text),
    Desc     = Table.AddColumn(Role, "Description", each if [RowType] = "ADDITIONAL" then [ColB] else null, type text),
    Unit     = Table.AddColumn(Desc, "UnitType", each
                  if List.Contains({"PSU_COUNT", "MA_PSU"}, [RowType]) then
                      (if [ColC] <> null then [ColC] else if [UnitLevel] = "L2" then "Level 2 PSU" else "Level 3 BDU")
                  else null, type text),
    Typed    = Table.TransformColumnTypes(Table.RemoveColumns(Unit, {"ColA", "ColB", "ColC", "RankText"}),
                  {{"NUM", type number}, {"TOD", type number}, {"XH", type number}})
in
    Typed
