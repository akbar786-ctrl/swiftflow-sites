// Query name: DimDate
let
    Start = #date(2028, 6, 1),
    Dates = List.Dates(Start, 61, #duration(1, 0, 0, 0)),
    T = Table.FromList(Dates, Splitter.SplitByNothing(), type table [Date = date]),
    DayName = Table.AddColumn(T, "Day", each Date.ToText([Date], "ddd dd MMM"), type text),
    Week = Table.AddColumn(DayName, "Week starting", each Date.StartOfWeek([Date], Day.Monday), type date)
in
    Week
