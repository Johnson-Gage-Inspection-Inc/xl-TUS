let
    // Try to get data with proper column names, fallback if columns don't exist yet
    Source = try
        let
            // Try to get the properly structured data
            RawSource = Excel.CurrentWorkbook(){[Name="UUTRangeTol"]}[Content],
            // Check if we have proper column names
            ColumnNames = Table.ColumnNames(RawSource),
            HasProperColumns = List.Contains(ColumnNames, "AssetId") and List.Contains(ColumnNames, "RangeText") and List.Contains(ColumnNames, "TolText"),
            
            ProcessedData = if HasProperColumns then
                // We have the expected columns, use them directly
                Table.SelectColumns(RawSource, {"AssetId", "RangeText", "TolText"})
            else
                // We don't have proper columns (probably during initialization), return empty table
                #table(
                    {"AssetId", "RangeText", "TolText"},
                    {{0, "Initializing...", "Initializing..."}}
                )
        in
            ProcessedData
    otherwise 
        // Fallback if table doesn't exist at all
        #table(
            {"AssetId", "RangeText", "TolText"},
            {{0, "Data unavailable", "Data unavailable"}}
        ),

    #"Grouped Rows" = Table.Group(Source, {"AssetId"}, {
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