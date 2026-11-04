// ============================================================================
// ROUTE B (fallback) - backup calculation workbooks
// BackupWorkbooks (load: yes) - one row per completed backup calculation
// workbook in folder pBackupFolder (create a Text parameter like
// pSubmissionsFolder). The workbook has already calculated everything with
// the legacy rules; Power BI only reads its output pages. Latest
// SubmissionNo per force is used.
// ============================================================================
let
    Files    = Folder.Files(pBackupFolder),
    Xlsx     = Table.SelectRows(Files, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")),
    Cover    = Table.AddColumn(Xlsx, "Cover", each
                   let Book = Excel.Workbook([Content], false, true),
                       Get  = (n as text) => try Book{[Item = n, Kind = "DefinedName"]}[Data]{0}[Column1] otherwise null
                   in  [Force = Get("cv_Force"), SubmissionNo = Get("cv_SubmissionNo"), TemplateVersion = Get("cv_TemplateVersion")]),
    Expanded = Table.ExpandRecordColumn(Cover, "Cover", {"Force", "SubmissionNo", "TemplateVersion"}),
    ForceKey = Table.AddColumn(Expanded, "ForceKey", each if [Force] = null then "UNKNOWN: " & [Name] else [Force], type text),
    No       = Table.AddColumn(ForceKey, "__No", each if [SubmissionNo] = null then 0 else [SubmissionNo], Int64.Type),
    Best     = Table.Group(No, {"ForceKey"}, {{"Keep", each Table.Max(_, {"__No", "Date modified"})[Name], type text}}),
    Current  = Table.AddColumn(No, "IsCurrent", each List.Contains(Best[Keep], [Name]), type logical),
    Result   = Table.SelectColumns(Current, {"Name", "Content", "Date modified", "ForceKey", "SubmissionNo", "TemplateVersion", "IsCurrent"})
in
    Result
