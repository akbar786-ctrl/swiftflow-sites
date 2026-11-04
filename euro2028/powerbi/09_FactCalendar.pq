// ============================================================================
// FactCalendar (load: yes) - Match Day demand and cost placed on real dates
// once RefFixture holds fixtures (date, host force, assumed risk level).
// Empty until then (D-01). Mirrors engine.expand_calendar.
//   AGREED  each fixture's MD-2..MD+1 rows are dated fixture date + offset and
//           coinciding dates are ADDED (D-02).
//   LEGACY  for each date, risk level and line only ONE relative day is used:
//           the first in the line's priority group
//             STD    MD > MD-1 > MD-2 > MD+1
//             NONMD  MD-1 > MD-2 > MD+1
//             MDONLY MD
//           Fixtures at different risk levels are added (legacy summed L+M+H).
// Kind = "Officer" (FactDeployed) or "Package" (FactMAPackage).
// ============================================================================
let
    Params    = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    Legacy    = Text.Upper(Text.From(Params[RuleSet])) = "LEGACY",
    Fixtures  = Table.SelectRows(RefFixture, each [MatchDate] <> null and [HostForce] <> null and [RiskCode] <> null),
    Officer   = Table.AddColumn(Table.SelectColumns(Table.SelectRows(FactDeployed, each [Collection] = "MD"),
                    {"ForceKey", "ResourceID", "LineKey", "RankCode", "SupplyCode", "DayGroup", "RiskCode", "DayCode",
                     "Officers", "ToDHours", "XHHours", "Cost"}), "Kind", each "Officer", type text),
    Package   = Table.AddColumn(Table.RenameColumns(Table.SelectColumns(Table.SelectRows(FactMAPackage, each [Collection] = "MD"),
                    {"ForceKey", "ResourceID", "LineKey", "DayGroup", "RiskCode", "DayCode", "Units", "Cost"}),
                    {{"Units", "Officers"}}), "Kind", each "Package", type text),
    Demand    = Table.Combine({Officer, Package}),
    Offsets   = Table.SelectColumns(DimRelativeDay, {"DayCode", "OffsetDays"}),
    Joined    = Table.ExpandTableColumn(Table.NestedJoin(Fixtures, {"HostForce", "RiskCode"}, Demand, {"ForceKey", "RiskCode"}, "D", JoinKind.Inner),
                    "D", {"ResourceID", "LineKey", "RankCode", "SupplyCode", "DayGroup", "DayCode", "Officers", "ToDHours", "XHHours", "Cost", "Kind"}),
    WithOff   = Table.ExpandTableColumn(Table.NestedJoin(Joined, {"DayCode"}, Offsets, {"DayCode"}, "O", JoinKind.LeftOuter), "O", {"OffsetDays"}),
    Dated     = Table.AddColumn(WithOff, "Date", each Date.AddDays(Date.From([MatchDate]), [OffsetDays]), type date),
    Priority  = Table.AddColumn(Dated, "Priority", each
                    let order = if [DayGroup] = "NONMD" then {"MD-1", "MD-2", "MD+1"}
                                else if [DayGroup] = "MDONLY" then {"MD"} else {"MD", "MD-1", "MD-2", "MD+1"},
                        p = List.PositionOf(order, [DayCode])
                    in  if p < 0 then 99 else p, Int64.Type),
    // LEGACY: the relative day chosen for a date is decided by which FIXTURES fall on the
    // surrounding dates (not by which rows hold values), so compute it per force x risk x date x group.
    Slots     = Table.Distinct(Table.SelectColumns(Table.SelectRows(Priority, each [Priority] < 99),
                    {"HostForce", "RiskCode", "Date", "DayGroup"})),
    FixDays   = Table.Distinct(Table.SelectColumns(Fixtures, {"HostForce", "RiskCode", "MatchDate"})),
    Chosen    = Table.AddColumn(Slots, "ChosenDay", each
                    let fx = List.Transform(Table.SelectRows(FixDays, (f) => f[HostForce] = [HostForce] and f[RiskCode] = [RiskCode])[MatchDate], Date.From),
                        order = if [DayGroup] = "NONMD" then {"MD-1", "MD-2", "MD+1"}
                                else if [DayGroup] = "MDONLY" then {"MD"} else {"MD", "MD-1", "MD-2", "MD+1"},
                        off = [#"MD-2" = -2, #"MD-1" = -1, MD = 0, #"MD+1" = 1],
                        hits = List.Select(order, (d) => List.Contains(fx, Date.AddDays([Date], -Record.Field(off, d))))
                    in  if List.IsEmpty(hits) then null else hits{0}, type text),
    Legacy1   = Table.ExpandTableColumn(Table.NestedJoin(Priority, {"HostForce", "RiskCode", "Date", "DayGroup"}, Chosen,
                    {"HostForce", "RiskCode", "Date", "DayGroup"}, "C", JoinKind.LeftOuter), "C", {"ChosenDay"}),
    Kept      = if Legacy then Table.SelectRows(Legacy1, each [DayCode] = [ChosenDay]) else Priority,
    Result    = Table.SelectColumns(Kept, {"FixtureID", "Venue", "HostForce", "RiskCode", "Date", "DayCode", "Kind",
                    "ResourceID", "LineKey", "RankCode", "SupplyCode", "Officers", "ToDHours", "XHHours", "Cost"})
in
    Result
