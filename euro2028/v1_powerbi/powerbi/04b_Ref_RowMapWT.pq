// Query name: Ref_RowMapWT   (do not load)
Table.TransformColumnTypes(Table.SelectRows(fnRefTable("tblRowMapWT"), each [Excel row] <> null), {
    {"Excel row", Int64.Type}, {"Section", type text}, {"Resource", type text}, {"Supplied by", type text},
    {"Line type", type text}, {"Rank", type text}, {"Rank from column C", type text}, {"Role from column C", type text},
    {"PSU group", Int64.Type}, {"Day rule", type text}, {"Rate supplied by", type text}, {"Rate rank", type text}})
