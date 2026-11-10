// Query name: Ref_Fixtures   (do not load)
let
    T = Table.SelectRows(fnRefTable("tblFixtures"), each [Date] <> null and [Risk] <> null and [Force] <> null)
in
    Table.TransformColumns(T, {
        {"Fixture", each if _ = null then "" else Text.From(_), type text},
        {"Date", each Date.From(_), type date},
        {"Risk", each Text.Trim(Text.From(_)), type text},
        {"Force", each Text.Trim(Text.From(_)), type text}})
