// Query name: DimForce   (LOAD)
let
    Names = List.Select(List.Distinct(List.Combine({Returns[Force], Ref_Fixtures[Force]})), each _ <> null and _ <> "")
in
    Table.Sort(Table.FromList(Names, Splitter.SplitByNothing(), type table [Force = text]), {{"Force", Order.Ascending}})
