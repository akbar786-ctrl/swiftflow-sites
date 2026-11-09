// Query name: DimForce
let
    Names = List.Distinct(List.Combine({MD_Fact[Force], WT_Fact[Force], Ref_Fixtures[Force]}))
in
    Table.Sort(Table.FromList(List.Select(Names, each _ <> null), Splitter.SplitByNothing(), type table [Force = text]), {{"Force", Order.Ascending}})
