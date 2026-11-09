// Query name: Calendar_Fact   (LOAD)
// Places each force's Match Day figures on real dates using the fixture list (Ref_Fixtures).
// For every fixture: MD-2 = date - 2, MD-1 = date - 1, MD = date, MD+1 = date + 1.
// When one force has more than one fixture-day on the same date at the same risk level, ONE day is used
// (original workbook rule): Match Day > MD-1 > MD-2 > MD+1; lines for days other than Match Day use
// MD-1 > MD-2 > MD+1.
let
    Offsets = #table(
        type table [Day = text, Offset = number, Priority = number, PriorityNonMD = nullable number],
        {{"MD-2", -2, 3, 2}, {"MD-1", -1, 2, 1}, {"MD", 0, 1, null}, {"MD+1", 1, 4, 3}}),
    WithOffsets = Table.AddColumn(Ref_Fixtures, "Rel", each Offsets),
    Expanded = Table.ExpandTableColumn(WithOffsets, "Rel", {"Day", "Offset", "Priority", "PriorityNonMD"}),
    Dated = Table.AddColumn(Expanded, "Calendar date", each Date.AddDays([Date], [Offset]), type date),
    Grouped = Table.Group(Dated, {"Force", "Calendar date", "Risk"}, {
        {"STD", each Table.Sort(_, {{"Priority", Order.Ascending}}){0}[Day], type text},
        {"NON_MD", each
            let NonMD = Table.SelectRows(_, each [PriorityNonMD] <> null)
            in if Table.RowCount(NonMD) = 0 then null else Table.Sort(NonMD, {{"PriorityNonMD", Order.Ascending}}){0}[Day],
            type nullable text},
        {"Fixtures", each Text.Combine(List.Sort(List.Distinct([Fixture])), ", "), type text}}),
    Std = Table.AddColumn(Table.RenameColumns(Table.SelectColumns(Grouped, {"Force", "Calendar date", "Risk", "STD", "Fixtures"}), {{"STD", "Day"}}), "Row type", each "STD", type text),
    NonMDRows = Table.SelectRows(Grouped, each [NON_MD] <> null),
    NonMD = Table.AddColumn(Table.RenameColumns(Table.SelectColumns(NonMDRows, {"Force", "Calendar date", "Risk", "NON_MD", "Fixtures"}), {{"NON_MD", "Day"}}), "Row type", each "NON_MD", type text),
    Lookup = Table.Combine({Std, NonMD}),
    Joined = Table.NestedJoin(MD_Fact, {"Force", "Risk", "Day", "Row type"}, Lookup, {"Force", "Risk", "Day", "Row type"}, "Cal", JoinKind.Inner),
    Result = Table.ExpandTableColumn(Joined, "Cal", {"Calendar date", "Fixtures"})
in
    Result
