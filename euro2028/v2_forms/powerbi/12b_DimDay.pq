// Query name: DimDay
#table(type table [Day = text, #"Day name" = text, #"Day order" = Int64.Type],
    {{"MD-2", "Match Day -2", 1}, {"MD-1", "Match Day -1", 2}, {"MD", "Match Day", 3}, {"MD+1", "Match Day +1", 4}})
