// ============================================================================
// Submissions  (load: yes - "DimSubmission")
// One row per force workbook in the submissions folder, with cover details,
// and IsCurrent = the latest SubmissionNo per force (older ones are kept for
// audit but excluded from all facts; reported as DQ16).
// ============================================================================
let
    Files      = Folder.Files(pSubmissionsFolder),
    Xlsx       = Table.SelectRows(Files, each Text.Lower([Extension]) = ".xlsx"
                                     and not Text.StartsWith([Name], "~$")
                                     and [Attributes]?[Hidden]? <> true),
    Cover      = Table.AddColumn(Xlsx, "Cover", each
                    let
                        Book  = Excel.Workbook([Content], false, true),
                        Get   = (n as text) => try Book{[Item = n, Kind = "DefinedName"]}[Data]{0}[Column1] otherwise null
                    in [
                        Force           = Get("cv_Force"),
                        ContactName     = Get("cv_ContactName"),
                        ContactEmail    = Get("cv_ContactEmail"),
                        SubmissionDate  = Get("cv_SubmissionDate"),
                        SubmissionNo    = Get("cv_SubmissionNo"),
                        TemplateVersion = Get("cv_TemplateVersion"),
                        Comments        = Get("cv_Comments")
                    ]),
    Expanded   = Table.ExpandRecordColumn(Cover, "Cover",
                    {"Force", "ContactName", "ContactEmail", "SubmissionDate", "SubmissionNo", "TemplateVersion", "Comments"}),
    Typed      = Table.TransformColumnTypes(Expanded, {{"SubmissionNo", Int64.Type}, {"TemplateVersion", type text},
                                                       {"Force", type text}}),
    ForceKey   = Table.AddColumn(Typed, "ForceKey", each if [Force] = null then "UNKNOWN: " & [Name] else [Force], type text),
    SubId      = Table.AddColumn(ForceKey, "SubmissionID",
                    each [ForceKey] & "#" & Text.From(if [SubmissionNo] = null then 0 else [SubmissionNo]) & "#" & [Name], type text),
    Latest     = Table.Group(SubId, {"ForceKey"}, {{"MaxNo", each List.Max(List.Transform([SubmissionNo], each if _ = null then 0 else _)), Int64.Type}}),
    Joined     = Table.NestedJoin(SubId, {"ForceKey"}, Latest, {"ForceKey"}, "L", JoinKind.LeftOuter),
    WithMax    = Table.ExpandTableColumn(Joined, "L", {"MaxNo"}),
    // ties on SubmissionNo: newest file modified date wins
    Ranked     = Table.AddColumn(WithMax, "IsCurrent", each (if [SubmissionNo] = null then 0 else [SubmissionNo]) = [MaxNo], type logical),
    TieBreak   = Table.Group(Table.SelectRows(Ranked, each [IsCurrent]), {"ForceKey"},
                    {{"Keep", each Table.Max(_, "Date modified")[SubmissionID], type text}}),
    Final      = Table.AddColumn(Ranked, "IsCurrentFinal", each List.Contains(TieBreak[Keep], [SubmissionID]), type logical),
    Result     = Table.RemoveColumns(Table.RenameColumns(Table.RemoveColumns(Final, {"IsCurrent", "MaxNo"}),
                                                         {{"IsCurrentFinal", "IsCurrent"}}),
                                     {"Extension", "Attributes", "Date accessed", "Date created"})
in
    Result
