// Query name: DimDate   (LOAD) - June and July 2028
let
    T = Table.FromList(List.Dates(#date(2028, 6, 1), 61, #duration(1, 0, 0, 0)), Splitter.SplitByNothing(), type table [Date = date]),
    D = Table.AddColumn(T, "Day label", each Date.ToText([Date], "ddd dd MMM"), type text)
in
    Table.AddColumn(D, "Week starting", each Date.StartOfWeek([Date], Day.Monday), type date)
