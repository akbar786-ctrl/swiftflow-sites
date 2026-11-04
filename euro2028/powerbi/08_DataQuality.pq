// ============================================================================
// DataQuality (load: yes) - one row per issue. Codes match engine/e28_engine.py:
//   DQ01 Cover field missing (force = Error, contact = Warning)
//   DQ02 Template version not accepted                       Error
//   DQ03 Structure changed (unknown / missing row keys)       Error
//   DQ04 Invalid value (text, negative, decimal Number)       Error   - excluded
//   DQ05 Hours above 24 (cell, or ToD + Extra on a line)      Error   - cell excluded
//   DQ06 Value typed into a not-applicable cell               Warning - excluded
//   DQ07 Number entered but no hours                          Warning
//   DQ08 Hours entered but no number                          Warning
//   DQ09 Rank not selected (not costed; legacy additional = Warning)
//   DQ10 Role profile not selected                            Warning
//   DQ11 Additional resource not described                    Warning
//   DQ12 Section not marked Complete / Not required           Warning
//   DQ13 PSU count with empty PSU configuration               Error
//   DQ14 No rate (rank or package)                            Warning
//   DQ16 Superseded submission                                Info
// ============================================================================
let
    Params     = Record.FromList(RefParameter[Value], RefParameter[Parameter]),
    Accepted   = Text.From(Params[TemplateVersion]),
    Legacy     = Text.Upper(Text.From(Params[RuleSet])) = "LEGACY",
    Mk         = (sub, force, coll, rowkey, code, sev, detail) =>
                    [SubmissionID = sub, ForceKey = force, Collection = coll, RowKey = rowkey, Code = code, Severity = sev, Detail = detail],
    Z          = (x) => if x = null then 0 else x,

    Cover      = Table.Combine(Table.TransformRows(Submissions, (s) => Table.FromRecords(List.RemoveNulls({
                    if not s[IsCurrent] then Mk(s[SubmissionID], s[ForceKey], null, null, "DQ16", "Info", "Superseded by a later submission") else null,
                    if s[IsCurrent] and s[Force] = null then Mk(s[SubmissionID], s[ForceKey], null, null, "DQ01", "Error", "Force not selected on Cover") else null,
                    if s[IsCurrent] and s[ContactName] = null then Mk(s[SubmissionID], s[ForceKey], null, null, "DQ01", "Warning", "Contact name missing") else null,
                    if s[IsCurrent] and s[ContactEmail] = null then Mk(s[SubmissionID], s[ForceKey], null, null, "DQ01", "Warning", "Contact email missing") else null,
                    if s[IsCurrent] and s[TemplateVersion] <> Accepted then Mk(s[SubmissionID], s[ForceKey], null, null, "DQ02", "Error",
                        "Template version " & (if s[TemplateVersion] = null then "(none)" else s[TemplateVersion]) & " not accepted") else null})))),

    Current    = Table.SelectRows(Submissions, each [IsCurrent]),
    Present    = Table.Combine(Table.TransformRows(Current, (s) => Table.AddColumn(Table.AddColumn(Table.Distinct(Table.SelectColumns(
                    Table.Combine({fnReadCollection(s[Content], "Match Day Resource"), fnReadCollection(s[Content], "Wider Tournament Resource")}),
                    {"RowKey"})), "SubmissionID", each s[SubmissionID]), "ForceKey", each s[ForceKey]))),
    Unknown    = Table.NestedJoin(Present, {"RowKey"}, DimResourceLine, {"RowKey"}, "X", JoinKind.LeftAnti),
    UnknownDQ  = Table.FromRecords(Table.TransformRows(Unknown, (r) => Mk(r[SubmissionID], r[ForceKey], Text.Start(r[RowKey], 2), r[RowKey], "DQ03", "Error", "Unknown row key - rows inserted or edited"))),
    Expected   = Table.ExpandTableColumn(Table.AddColumn(Table.SelectColumns(Current, {"SubmissionID", "ForceKey"}), "K",
                    each Table.SelectColumns(DimResourceLine, {"RowKey"})), "K", {"RowKey"}),
    Missing    = Table.NestedJoin(Expected, {"SubmissionID", "RowKey"}, Present, {"SubmissionID", "RowKey"}, "X", JoinKind.LeftAnti),
    MissingDQ  = Table.FromRecords(Table.TransformRows(Missing, (r) => Mk(r[SubmissionID], r[ForceKey], Text.Start(r[RowKey], 2), r[RowKey], "DQ03", "Error", "Expected row missing - rows deleted"))),

    CellIssues = Table.SelectRows(FactInputCell, each List.Contains({"INVALID_TEXT", "INVALID_NEGATIVE", "INVALID_DECIMAL", "INVALID_HOURS", "NA_WITH_VALUE"}, [ValueState])),
    CellDQ     = Table.FromRecords(Table.TransformRows(CellIssues, (r) => Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey],
                    if r[ValueState] = "INVALID_HOURS" then "DQ05" else if r[ValueState] = "NA_WITH_VALUE" then "DQ06" else "DQ04",
                    if r[ValueState] = "NA_WITH_VALUE" then "Warning" else "Error",
                    r[RiskCode] & " " & r[DayCode] & " " & r[MeasureCode] & ": " &
                    (if r[ValueState] = "INVALID_TEXT" then "non-numeric value '" & Text.From(r[RawValue]) & "'"
                     else if r[ValueState] = "INVALID_NEGATIVE" then "negative value"
                     else if r[ValueState] = "INVALID_DECIMAL" then "number required must be a whole number"
                     else if r[ValueState] = "INVALID_HOURS" then "hours above 24"
                     else "value in a not-applicable cell") & " - excluded"))),

    Officer    = Table.SelectRows(FactInput, each List.Contains({"OFFICER", "RANK_SELECT", "ROLE", "ADDITIONAL"}, [RowType])),
    LineDQ     = Table.Combine(Table.TransformRows(Officer, (r) =>
                    let has = Z(r[NUM]) > 0, ctx = r[RiskCode] & " " & r[DayCode] & ": ",
                        legacyAdd = Legacy and r[RowType] = "ADDITIONAL"
                    in Table.FromRecords(List.RemoveNulls({
                        if has and Z(r[TOD]) = 0 and Z(r[XH]) = 0 then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ07", "Warning", ctx & "number entered but no hours") else null,
                        if has and Z(r[TOD]) + Z(r[XH]) > 24 then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ05", "Error", ctx & "Normal ToD + Extra Hours above 24") else null,
                        if has and r[RankCode] = null then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ09",
                            if legacyAdd then "Warning" else "Error", ctx & "rank not selected" & (if legacyAdd then "" else " - not costed")) else null,
                        if has and r[RowType] = "ROLE" and r[Role] = null then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ10", "Warning", ctx & "role profile not selected") else null,
                        if has and r[RowType] = "ADDITIONAL" and r[Description] = null then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ11", "Warning", ctx & "additional resource not described") else null,
                        if not has and (Z(r[TOD]) > 0 or Z(r[XH]) > 0) then Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ08", "Warning", ctx & "hours entered but no number") else null})))),

    Sections   = Table.Distinct(Table.SelectColumns(FactInputCell, {"SubmissionID", "ForceKey", "Collection", "SectionKey", "SectionStatus"})),
    SectionDQ  = Table.FromRecords(Table.TransformRows(Table.SelectRows(Sections, each not List.Contains({"Complete", "Not required"}, [SectionStatus])),
                    (r) => Mk(r[SubmissionID], r[ForceKey], r[Collection], "SECTION|" & r[SectionKey], "DQ12", "Warning", "Section " & r[SectionKey] & " status: " & r[SectionStatus]))),

    Counts     = Table.SelectRows(FactInput, each [RowType] = "PSU_COUNT" and Z([NUM]) > 0),
    Expanded   = Table.Distinct(Table.SelectColumns(Table.SelectRows(FactDeployed, each [Source] = "PSU expansion"), {"SubmissionID", "RowKey", "RiskCode", "DayCode"})),
    NoCfg      = Table.NestedJoin(Counts, {"SubmissionID", "RowKey", "RiskCode", "DayCode"}, Expanded, {"SubmissionID", "RowKey", "RiskCode", "DayCode"}, "X", JoinKind.LeftAnti),
    NoCfgDQ    = Table.FromRecords(Table.TransformRows(NoCfg, (r) => Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ13", "Error",
                    r[RiskCode] & " " & r[DayCode] & ": " & Text.From(r[NUM]) & " units entered but the PSU configuration is empty"))),

    NoRate     = Table.SelectRows(FactDeployed, each [RankCode] <> null and [OTRate] = null),
    NoRateDQ   = Table.FromRecords(Table.TransformRows(NoRate, (r) => Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ14", "Warning", "No rate for " & r[RateKey] & " - not costed"))),
    NoPkg      = Table.Distinct(Table.SelectColumns(Table.SelectRows(FactMAPackage, each [PackageRate] = null), {"SubmissionID", "ForceKey", "Collection", "RowKey", "UnitType"})),
    NoPkgDQ    = Table.FromRecords(Table.TransformRows(NoPkg, (r) => Mk(r[SubmissionID], r[ForceKey], r[Collection], r[RowKey], "DQ14", "Warning", "No mutual-aid package rate for '" & r[UnitType] & "' - not costed (OI-03)"))),

    All        = Table.Combine({Cover, UnknownDQ, MissingDQ, CellDQ, LineDQ, SectionDQ, NoCfgDQ, NoRateDQ, NoPkgDQ}),
    Typed      = Table.TransformColumnTypes(All, {{"SubmissionID", type text}, {"ForceKey", type text}, {"Collection", type text},
                    {"RowKey", type text}, {"Code", type text}, {"Severity", type text}, {"Detail", type text}}),
    Indexed    = Table.AddIndexColumn(Typed, "IssueID", 1, 1, Int64.Type)
in
    Indexed
