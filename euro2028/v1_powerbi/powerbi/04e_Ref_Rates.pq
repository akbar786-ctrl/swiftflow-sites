// Query name: Ref_Rates   (do not load)
Table.SelectRows(Table.TransformColumnTypes(fnRefTable("tblRates"), {
    {"Supplied by", type text}, {"Rank", type text}, {"Hourly rate", type number}}), each [Rank] <> null)
