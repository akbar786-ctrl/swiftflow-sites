// Query name: Ref_WTDays   (do not load)
Table.SelectRows(Table.TransformColumnTypes(fnRefTable("tblWTDays"), {{"Risk", type text}, {"Days", type number}}), each [Risk] <> null)
