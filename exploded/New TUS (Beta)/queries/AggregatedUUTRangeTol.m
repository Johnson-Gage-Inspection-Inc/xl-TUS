let
    Source = Excel.CurrentWorkbook(){[Name="UUTRangeTol"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"AssetId", Int64.Type}, {"RangeMin", Int64.Type}, {"RangeMax", Int64.Type}, {"TolMin", type number}, {"TolMax", type number}, {"TolResolution", type number}, {"Unit", type text}, {"Flags", type text}, {"TolMinText", type text}, {"TolMaxText", type text}, {"TolText", type text}, {"RangeMaxText", type text}, {"RangeText", type text}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"AssetId", "RangeText", "TolText"}),

    #"Grouped Rows" = Table.Group(#"Removed Other Columns", {"AssetId"}, {
        {"Ranges", each Text.Combine(List.Transform([RangeText], Text.From), ", "), type text},
        {"Tolerances", each 
            let
                allTols = List.Transform([TolText], Text.From),
                distinctTols = List.Distinct(allTols)
            in
                if List.Count(distinctTols) = 1 then
                    distinctTols{0}
                else
                    Text.Combine(allTols, " / ")
        , type text}
    })
in
    #"Grouped Rows"