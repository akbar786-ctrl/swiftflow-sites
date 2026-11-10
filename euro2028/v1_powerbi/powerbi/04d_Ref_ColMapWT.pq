// Query name: Ref_ColMapWT   (do not load)
Table.TransformColumnTypes(Table.SelectRows(fnRefTable("tblColMapWT"), each [Column number] <> null), {
    {"Column number", Int64.Type}, {"Column", type text}, {"Risk", type text}, {"Day", type text}, {"Measure", type text}})
