// Query name: DataQuality   (LOAD)
// Things to chase with forces before trusting the figures. Empty = nothing to chase.
let
    NotForms = Table.SelectRows(
        Table.AddColumn(Returns, "Ok", each (try fnFormLong([Content], [Name], "tblMatchDay"))[HasError] = false),
        each not [Ok]),
    NotFormIssues = Table.FromRecords(List.Transform(NotForms[Name], each [Force = null, #"Source file" = _, Page = "-",
        Line = "-", Issue = "File is not a EURO 2028 collection form (no MD Collection table)"]),
        type table [Force = nullable text, #"Source file" = text, Page = text, Line = text, Issue = text]),
    LongAll = Table.Combine({
        Table.AddColumn(MD_Long, "Page", each "MD Collection", type text),
        Table.AddColumn(WT_Long, "Page", each "WT Collection", type text)}),
    LineText = (r as record) as text => Text.Combine(List.Select({r[Section], r[Resource], r[#"Supplied by"], r[#"Role / Skill"], r[Rank]}, each _ <> null and _ <> "" and _ <> "-"), " / "),
    NotNumber = Table.SelectRows(LongAll, each [Raw] <> null and [Value] = null),
    NotNumberIssues = Table.AddColumn(NotNumber, "Issue", each "Not a number in '" & [Risk] & " " & [Day] & " " & [Measure] & "': " & [Raw], type text),
    Filled = Table.SelectRows(LongAll, each [Value] <> null and [Value] <> 0),
    Lines = Table.Distinct(Table.SelectColumns(Filled, {"Force", "Source file", "Page", "Form row", "Section", "Resource", "Supplied by", "Role / Skill", "Rank"})),
    MissingIssues = Table.SelectRows(Table.AddColumn(Lines, "Issue", each
        if [Rank] = null or [Rank] = "" then "Numbers entered but no Rank chosen"
        else if [Resource] = null or [Resource] = "" then "Numbers entered but no Resource given"
        else if [#"Supplied by"] = null or [#"Supplied by"] = "" then "Numbers entered but Supplied by not chosen"
        else if [#"Role / Skill"] = null or [#"Role / Skill"] = "" then "Numbers entered but no Role / Skill (or unit type) chosen"
        else null, type nullable text), each [Issue] <> null),
    Hours = Table.SelectRows(Filled, each [Measure] <> "Officers" and [Value] > 24),
    HoursIssues = Table.AddColumn(Hours, "Issue", each "More than 24 hours per officer in '" & [Risk] & " " & [Day] & " " & [Measure] & "'", type text),
    NoRate = Table.SelectRows(Table.Combine({
        Table.AddColumn(MD_Fact, "Page", each "MD Collection", type text),
        Table.AddColumn(WT_Fact, "Page", each "WT Collection", type text)}), each [Officers] > 0 and [Hourly rate] = null),
    NoRateIssues = Table.AddColumn(Table.Distinct(Table.SelectColumns(NoRate, {"Force", "Source file", "Page", "Form row", "Section", "Resource", "Supplied by", "Role / Skill", "Rank"})),
        "Issue", each "No hourly rate for " & ([#"Supplied by"] ?? "?") & " / " & ([Rank] ?? "?") & " - add it to tblRates", type text),
    Shape = (t as table) as table => Table.SelectColumns(Table.AddColumn(t, "Line", each "Row " & Text.From([#"Form row"]) & ": " & LineText(_), type text), {"Force", "Source file", "Page", "Line", "Issue"}),
    FixtureForces = List.Distinct(Ref_Fixtures[Force]),
    KnownForces = List.Distinct(MD_Fact[Force]),
    UnknownFixtureForces = List.Select(FixtureForces, each not List.Contains(KnownForces, _)),
    FixtureIssues = Table.FromRecords(List.Transform(UnknownFixtureForces, each [Force = _, #"Source file" = "Reference data",
        Page = "Fixtures", Line = "-", Issue = "Force on the fixture list has no Match Day figures (check the spelling matches the form)"]),
        type table [Force = nullable text, #"Source file" = text, Page = text, Line = text, Issue = text]),
    All = Table.Combine({NotFormIssues, Shape(NotNumberIssues), Shape(MissingIssues), Shape(HoursIssues), Shape(NoRateIssues), FixtureIssues})
in
    All
