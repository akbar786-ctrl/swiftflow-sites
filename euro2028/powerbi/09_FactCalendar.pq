// ============================================================================
// FactCalendar (load: yes) - places Match Day demand on real dates once
// RefFixture is populated. Empty until then (D-01: fixtures not set).
// Mirrors engine.expand_calendar:
//   * each fixture's host force and assumed risk level select that force's
//     demand rows for that risk;
//   * each row is dated fixture date + relative-day offset;
//   * where dates coincide (e.g. MD+1 of one fixture and MD-2 of the next)
//     rows are ADDED (Decision D-02) - there is no precedence rule.
// ============================================================================
let
    Fixtures  = Table.SelectRows(RefFixture, each [MatchDate] <> null and [HostForce] <> null and [RiskCode] <> null),
    MD        = Table.SelectRows(FactDeployed, each [Collection] = "MD"),
    Offsets   = Table.SelectColumns(DimRelativeDay, {"DayCode", "OffsetDays"}),
    Joined    = Table.NestedJoin(Fixtures, {"HostForce", "RiskCode"}, MD, {"ForceKey", "RiskCode"}, "D", JoinKind.Inner),
    Expanded  = Table.ExpandTableColumn(Joined, "D",
                    {"DeployedID", "DayCode", "ResourceID", "RankCode", "SupplyCode", "RateKey", "Officers", "ToDHours", "XHHours"}),
    WithOff   = Table.ExpandTableColumn(Table.NestedJoin(Expanded, {"DayCode"}, Offsets, {"DayCode"}, "O", JoinKind.LeftOuter), "O", {"OffsetDays"}),
    Dated     = Table.AddColumn(WithOff, "Date", each Date.AddDays(Date.From([MatchDate]), [OffsetDays]), type date),
    Result    = Table.SelectColumns(Dated, {"FixtureID", "Venue", "HostForce", "RiskCode", "Date", "DayCode", "DeployedID",
                    "ResourceID", "RankCode", "SupplyCode", "RateKey", "Officers", "ToDHours", "XHHours"})
in
    Result
