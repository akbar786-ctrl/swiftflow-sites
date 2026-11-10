// Query name: fnBuildFact   (function - do not load)
// Turns Lines into the costed fact table, exactly as the original collection workbook:
//   officer lines      Officers = Number;  cost = Number x Extra Hours x hourly rate
//   PSU / Serial lines units = Number; force supplied officers = units x officers per PSU (configuration rows),
//                      cost = officers x configuration Extra Hours x rate; mutual aid = units x package
// Hourly rate = Ref_Rates for (Rate supplied by, Rate rank) from the row map - rank from column C where the row has
// a rank drop-down.
(Lines as table) as table =>
let
    NormRank = (x as any) as text =>
        let t = if x = null then "" else Text.Trim(Text.Replace(Text.From(x), "#(00A0)", " "))
        in if t = "Sergeants" then "Sergeant" else if t = "Constables" then "Constable" else if t = "Supt" then "Superintendent" else t,
    CText = (x as any) as text => if x = null then "" else Text.Trim(Text.From(x)),
    AddRate = (t as table, supCol as text, rankCol as text) as table =>
        let
            J = Table.NestedJoin(t, {supCol, rankCol}, Ref_Rates, {"Supplied by", "Rank"}, "R", JoinKind.LeftOuter),
            A = Table.AddColumn(J, "Hourly rate", each if Table.RowCount([R]) > 0 then [R]{0}[Hourly rate] else 0, type number)
        in Table.RemoveColumns(A, {"R"}),
    Cols = {"Force", "Source file", "Excel row", "Section", "Resource", "Supplied by", "Line type", "Rank", "Role / Skill",
            "Risk", "Day", "Day rule", "Officers", "PSUs / Serials", "Hourly rate", "Extra hours total", "Normal hours total",
            "Overtime cost", "Normal duty cost", "Mutual aid PSU cost"},
    // ---- officer, role and additional lines
    Off0 = Table.SelectRows(Lines, each List.Contains({"Officer", "Role", "Additional"}, [Line type])),
    Off1 = Table.AddColumn(Off0, "Rank used", each if [Rank from column C] = "Yes" then NormRank([Column C]) else ([Rank] ?? ""), type text),
    Off2 = Table.AddColumn(Off1, "Role used", each if [Role from column C] = "Yes" then CText([Column C]) else "", type text),
    Off3 = Table.AddColumn(Off2, "Resource used", each if [Line type] = "Additional" and CText([Column B]) <> "" then CText([Column B]) else [Resource], type text),
    Off4 = Table.AddColumn(Off3, "Rate rank used", each if ([Rate rank] ?? "") = "" then [Rank used] else [Rate rank], type text),
    Off5 = AddRate(Off4, "Rate supplied by", "Rate rank used"),
    Off6 = Table.AddColumn(Off5, "Officers", each [Number], type number),
    Off7 = Table.AddColumn(Off6, "PSUs / Serials", each 0, type number),
    Off8 = Table.AddColumn(Off7, "Extra hours total", each [Number] * [Extra Hours], type number),
    Off9 = Table.AddColumn(Off8, "Normal hours total", each [Number] * [Normal ToD], type number),
    Off10 = Table.AddColumn(Off9, "Overtime cost", each [Number] * [Extra Hours] * [Hourly rate], type number),
    Off11 = Table.AddColumn(Off10, "Normal duty cost", each [Number] * [Normal ToD] * [Hourly rate], type number),
    Off12 = Table.AddColumn(Off11, "Mutual aid PSU cost", each 0, type number),
    OffOut = Table.SelectColumns(Table.RenameColumns(Table.RemoveColumns(Off12, {"Rank", "Resource"}),
        {{"Rank used", "Rank"}, {"Role used", "Role / Skill"}, {"Resource used", "Resource"}}), Cols),
    // ---- PSU / Serial unit lines (force supplied and mutual aid)
    Cnt0 = Table.SelectRows(Lines, each [Line type] = "PSU count" or [Line type] = "MA PSU count"),
    Cnt1 = Table.AddColumn(Cnt0, "Role / Skill", each CText([Column C]), type text),
    Units1 = Table.TransformColumns(Table.AddColumn(Cnt1, "Officers", each 0, type number), {{"Line type", each "PSU units", type text}}),
    Units2 = Table.AddColumn(Units1, "PSUs / Serials", each [Number], type number),
    Units3 = Table.AddColumn(Units2, "Hourly rate", each 0, type number),
    Units4 = Table.AddColumn(Units3, "Extra hours total", each 0, type number),
    Units5 = Table.AddColumn(Units4, "Normal hours total", each 0, type number),
    Units6 = Table.AddColumn(Units5, "Overtime cost", each 0, type number),
    Units7 = Table.AddColumn(Units6, "Normal duty cost", each 0, type number),
    Units8 = Table.AddColumn(Units7, "Mutual aid PSU cost", each if [#"Supplied by"] = "Mutual aid" then [Number] * Ref_PSUPackage else 0, type number),
    UnitsOut = Table.SelectColumns(Units8, Cols),
    // ---- force supplied PSU officers = units x officers per PSU from the configuration rows
    Cfg = Table.SelectColumns(Table.SelectRows(Lines, each [Line type] = "PSU config"),
        {"Source file", "PSU group", "Risk", "Day", "Rank", "Number", "Normal ToD", "Extra Hours", "Rate supplied by", "Rate rank"}),
    FS0 = Table.SelectRows(Cnt1, each [Line type] = "PSU count"),
    FS1 = Table.NestedJoin(Table.RemoveColumns(FS0, {"Rank", "Rate supplied by", "Rate rank"}), {"Source file", "PSU group", "Risk", "Day"},
        Cfg, {"Source file", "PSU group", "Risk", "Day"}, "Cfg", JoinKind.Inner),
    FS2 = Table.ExpandTableColumn(FS1, "Cfg", {"Rank", "Number", "Normal ToD", "Extra Hours", "Rate supplied by", "Rate rank"},
        {"Rank", "Per PSU", "Cfg Normal ToD", "Cfg Extra Hours", "Rate supplied by", "Rate rank"}),
    FS3 = AddRate(FS2, "Rate supplied by", "Rate rank"),
    FS4 = Table.TransformColumns(Table.AddColumn(FS3, "Officers", each [Number] * [Per PSU], type number), {{"Line type", each "PSU officers", type text}}),
    FS5 = Table.AddColumn(FS4, "PSUs / Serials", each 0, type number),
    FS6 = Table.AddColumn(FS5, "Extra hours total", each [Officers] * [Cfg Extra Hours], type number),
    FS7 = Table.AddColumn(FS6, "Normal hours total", each [Officers] * [Cfg Normal ToD], type number),
    FS8 = Table.AddColumn(FS7, "Overtime cost", each [Officers] * [Cfg Extra Hours] * [Hourly rate], type number),
    FS9 = Table.AddColumn(FS8, "Normal duty cost", each [Officers] * [Cfg Normal ToD] * [Hourly rate], type number),
    FS10 = Table.AddColumn(FS9, "Mutual aid PSU cost", each 0, type number),
    FSOut = Table.SelectColumns(FS10, Cols),
    All = Table.Combine({OffOut, UnitsOut, FSOut}),
    Typed = Table.TransformColumnTypes(All, {{"Force", type text}, {"Source file", type text}, {"Excel row", Int64.Type},
        {"Section", type text}, {"Resource", type text}, {"Supplied by", type text}, {"Line type", type text}, {"Rank", type text},
        {"Role / Skill", type text}, {"Risk", type text}, {"Day", type text}, {"Day rule", type text}})
in
    Typed
