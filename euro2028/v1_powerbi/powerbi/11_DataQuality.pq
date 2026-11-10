// Query name: DataQuality   (LOAD) - things to chase with forces before trusting the figures (empty = clean)
let
    Issue = (force as text, file as text, page as text, row as any, issue as text) as record =>
        [Force = force, #"Source file" = file, Page = page, #"Excel row" = row, Issue = issue],
    // files that are not the risk-first EURO 2028 template
    Checked = Table.AddColumn(Returns, "Problem", each
        let
            Book = try Excel.Workbook([Content], false, true) otherwise null,
            Names = if Book = null then {} else Table.SelectRows(Book, each [Kind] = "DefinedName")[Name],
            Grid = if List.Contains(Names, "MD_Data") then Table.SelectRows(Book, each [Kind] = "DefinedName" and [Name] = "MD_Data"){0}[Data] else null
        in
            if Book = null then "File could not be opened as Excel"
            else if not List.Contains(Names, "MD_Data") or not List.Contains(Names, "WT_Data") then "Not the EURO 2028 Power BI template (named ranges MD_Data / WT_Data missing)"
            else if Text.From(Grid{2}[Column4] ?? "") <> "Low" or Text.From(Grid{2}[Column40] ?? "") <> "Super High" then "Old layout: Match Day Resource row 3 should read Low / Medium / High / Super High"
            else null),
    FileIssues = List.Transform(Table.ToRecords(Table.SelectRows(Checked, each [Problem] <> null)),
        each Issue([Force], [Name], "-", null, [Problem])),
    Both = Table.Combine({
        Table.AddColumn(MD_Lines, "Page", each "Match Day Resource", type text),
        Table.AddColumn(WT_Lines, "Page", each "Wider Tournament Resource", type text)}),
    LineIssues = Table.SelectRows(Table.AddColumn(Both, "Issue", each
        if [Rank from column C] = "Yes" and Text.Trim(Text.From([Column C] ?? "")) = "" then "Numbers entered but no rank chosen in column C"
        else if [Line type] = "Role" and Text.Trim(Text.From([Column C] ?? "")) = "" then "Numbers entered but no role chosen in column C"
        else if [Line type] = "Additional" and Text.Trim(Text.From([Column B] ?? "")) = "" then "Additional resource numbers entered but no description in column B"
        else if [Number] = 0 and [Extra Hours] > 0 then "Hours entered but no number of officers"
        else if [Extra Hours] > 24 or [Normal ToD] > 24 then "More than 24 hours per officer"
        else null, type nullable text), each [Issue] <> null),
    LineRecords = List.Transform(Table.ToRecords(LineIssues),
        each Issue([Force], [Source file], [Page], [Excel row], [Issue] & " (" & [Risk] & " " & [Day] & ")")),
    PSU = Table.SelectRows(Table.Combine({MD_Fact, WT_Fact}), each [Line type] = "PSU units" and [#"Supplied by"] = "Force supplied"),
    PSUOfficers = Table.SelectRows(Table.Combine({MD_Fact, WT_Fact}), each [Line type] = "PSU officers"),
    NoConfig = Table.SelectRows(PSU, (u) => Table.RowCount(Table.SelectRows(PSUOfficers, each [Source file] = u[Source file]
        and [Resource] = u[Resource] and [Risk] = u[Risk] and [Day] = u[Day])) = 0),
    NoConfigRecords = List.Transform(Table.ToRecords(NoConfig), each Issue([Force], [Source file], "-", [Excel row],
        "PSUs / Serials entered but the PSU / Serial configuration (officers per PSU) is blank (" & [Risk] & " " & [Day] & ")")),
    KnownForces = List.Distinct(MD_Fact[Force]),
    FixtureRecords = List.Transform(List.Select(List.Distinct(Ref_Fixtures[Force]), each not List.Contains(KnownForces, _)),
        each Issue(_, "Reference data", "Fixtures", null, "Force on the fixture list has no Match Day figures - check the spelling matches the file name")),
    All = FileIssues & LineRecords & NoConfigRecords & FixtureRecords
in
    Table.FromRecords(All, type table [Force = nullable text, #"Source file" = nullable text, Page = nullable text,
        #"Excel row" = nullable number, Issue = text])
