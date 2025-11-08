let
    // Try to get fresh data from API
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "wire-offsets/"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column2" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"correction_factor", "nominal_temp", "traceability_no"}, {"correction_factor", "nominal_temp", "traceability_no"}),
            #"Renamed Columns" = Table.RenameColumns(#"Expanded Column2",{{"correction_factor", "Offset"}, {"nominal_temp", "Temp"}}),
            #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Offset", type number}, {"Temp", type number}, {"traceability_no", type text}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try Excel.CurrentWorkbook(){[Name="AllWireOffsets"]}[Content] otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"traceability_no", "Temp", "Offset"},
                    {{"Data unavailable: Unable to retrieve data from API or cached workbook. Please check your network connection and try again.", 0.0, 0.0}}
                ),
    #"Reordered Columns" = Table.ReorderColumns(FinalData,{"traceability_no", "Temp", "Offset"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"traceability_no", Order.Ascending}, {"Temp", Order.Ascending}})
in
    #"Sorted Rows"