// Query name: Ref_ColMapMD   (do not load)
Table.TransformColumnTypes(Table.SelectRows(fnRefTable("tblColMapMD"), each [Column number] <> null), {
    {"Column number", Int64.Type}, {"Column", type text}, {"Risk", type text}, {"Day", type text}, {"Measure", type text}})
