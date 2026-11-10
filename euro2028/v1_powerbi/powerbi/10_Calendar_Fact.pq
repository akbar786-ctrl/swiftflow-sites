// Query name: Calendar_Fact   (LOAD)
// Places MD_Fact on real dates using Ref_Fixtures, with the original workbook's rule: on each date, for each force
// and risk level, ONE day of the fixture cycle is used -
//   ordinary lines            Match Day > MD-1 > MD-2 > MD+1
//   non-matchday PSU lines    MD-1 > MD-2 > MD+1
//   matchday-only PSU lines   Match Day only
let
    Offsets = #table(type table [Day = text, Offset = number], {{"MD-2", -2}, {"MD-1", -1}, {"MD", 0}, {"MD+1", 1}}),
    Rules = #table(type table [#"Day rule" = text, Day = text, Priority = number], {
        {"STD", "MD", 1}, {"STD", "MD-1", 2}, {"STD", "MD-2", 3}, {"STD", "MD+1", 4},
        {"NONMD", "MD-1", 1}, {"NONMD", "MD-2", 2}, {"NONMD", "MD+1", 3},
        {"MDONLY", "MD", 1}}),
    WithOff = Table.ExpandTableColumn(Table.AddColumn(Ref_Fixtures, "O", each Offsets), "O", {"Day", "Offset"}),
    Dated = Table.AddColumn(WithOff, "Calendar date", each Date.AddDays([Date], [Offset]), type date),
    WithRules = Table.NestedJoin(Dated, {"Day"}, Rules, {"Day"}, "R", JoinKind.Inner),
    Exp = Table.ExpandTableColumn(WithRules, "R", {"Day rule", "Priority"}),
    Best = Table.Group(Exp, {"Force", "Calendar date", "Risk", "Day rule"}, {
        {"Day", each Table.Sort(_, {{"Priority", Order.Ascending}}){0}[Day], type text},
        {"Fixtures", each Text.Combine(List.Sort(List.Distinct([Fixture])), ", "), type text}}),
    Joined = Table.NestedJoin(MD_Fact, {"Force", "Risk", "Day", "Day rule"}, Best, {"Force", "Risk", "Day", "Day rule"}, "Cal", JoinKind.Inner)
in
    Table.ExpandTableColumn(Joined, "Cal", {"Calendar date", "Fixtures"})
