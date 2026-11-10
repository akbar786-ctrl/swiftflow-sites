// Query name: Ref_PSUPackage   (do not load) - a single number: mutual aid PSU / Serial cost per unit per day
let
    T = Table.TransformColumnTypes(fnRefTable("tblSettings"), {{"Key", type text}, {"Value", type number}})
in
    T{[Key = "MA_PSU_PACKAGE"]}[Value]
