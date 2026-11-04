// ============================================================================
// Submissions (load: yes) - one row per workbook in the submissions folder
// (force collection templates; backup calculation workbooks are read by
// 10_RouteB_*.pq instead). IsCurrent = latest SubmissionNo per force; ties go
// to the most recently modified file. Older files are kept for audit (DQ16).
// ============================================================================
let
    Files     = Folder.Files(pSubmissionsFolder),
    Xlsx      = Table.SelectRows(Files, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),
    Cover     = Table.AddColumn(Xlsx, "Cover", each
                    let
                        Book = Excel.Workbook([Content], false, true),
                        Get  = (n as text) => try Book{[Item = n, Kind = "DefinedName"]}[Data]{0}[Column1] otherwise null
                    in [Force = Get("cv_Force"), ContactName = Get("cv_ContactName"), ContactEmail = Get("cv_ContactEmail"),
                        SubmissionDate = Get("cv_SubmissionDate"), SubmissionNo = Get("cv_SubmissionNo"),
                        TemplateVersion = Get("cv_TemplateVersion"), Comments = Get("cv_Comments")]),
    Expanded  = Table.ExpandRecordColumn(Cover, "Cover", {"Force", "ContactName", "ContactEmail", "SubmissionDate",
                                                          "SubmissionNo", "TemplateVersion", "Comments"}),
    Typed     = Table.TransformColumnTypes(Expanded, {{"SubmissionNo", Int64.Type}, {"TemplateVersion", type text}, {"Force", type text}}),
    ForceKey  = Table.AddColumn(Typed, "ForceKey", each if [Force] = null then "UNKNOWN: " & [Name] else [Force], type text),
    SubId     = Table.AddColumn(ForceKey, "SubmissionID", each [ForceKey] & "#" &
                    Text.From(if [SubmissionNo] = null then 0 else [SubmissionNo]) & "#" & [Name], type text),
    NoOrZero  = Table.AddColumn(SubId, "__No", each if [SubmissionNo] = null then 0 else [SubmissionNo], Int64.Type),
    Best      = Table.Group(NoOrZero, {"ForceKey"}, {{"Keep", each Table.Max(_, {"__No", "Date modified"})[SubmissionID], type text}}),
    Current   = Table.AddColumn(NoOrZero, "IsCurrent", each List.Contains(Best[Keep], [SubmissionID]), type logical),
    Result    = Table.RemoveColumns(Current, {"__No", "Extension", "Attributes", "Date accessed", "Date created"})
in
    Result
