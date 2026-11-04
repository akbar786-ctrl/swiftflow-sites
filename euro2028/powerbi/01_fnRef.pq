// ============================================================================
// fnRef  (function)  - returns a named Excel table from the reference workbook,
// with blank rows removed. Every reference query is simply  = fnRef("DimRank")
// ============================================================================
(tableName as text) as table =>
let
    Source   = Excel.Workbook(File.Contents(pReferenceFile), true, true),
    Tbl      = Source{[Item = tableName, Kind = "Table"]}[Data],
    NonBlank = Table.SelectRows(Tbl, each List.NonNullCount(Record.FieldValues(_)) > 0)
in
    NonBlank

// ----------------------------------------------------------------------------
// Reference queries to create (one blank query each, "Enable load" as noted):
//   DimResource       = fnRef("DimResource")            load
//   DimResourceLine   = fnRef("DimResourceLine")        load
//   DimRank           = fnRef("DimRank")                load
//   DimRelativeDay    = fnRef("DimRelativeDay")         load
//   DimRoleProfile    = fnRef("DimRoleProfile")         load
//   DimForce          = fnRef("DimForce")               load
//   RefParameter      = fnRef("RefParameter")           no load (used by queries)
//   RefBaseRate       = fnRef("RefBaseRate")            no load
//   RefUpliftStep     = fnRef("RefUpliftStep")          no load
//   RefZeroCost       = fnRef("RefZeroCost")            no load
//   RefMAPackage      = fnRef("RefMAPackage")           no load
//   MapOldToNew       = fnRef("MapOldToNew")            load
//   LogDecision       = fnRef("LogDecision")            load
//   LogOpenItem       = fnRef("LogOpenItem")            load
//   RefEquipmentRatio = fnRef("RefEquipmentRatio")      load
//   RefScenarioUplift = fnRef("RefScenarioUplift")      load
//   RefCapacity       = fnRef("RefCapacity")            load
//   RefFixture        = fnRef("RefFixture")             load
//   RefFixturePlan    = fnRef("RefFixturePlan")         load
//   DimRisk           -> see 07_DimRate.pq (adds WT days)
// ----------------------------------------------------------------------------
