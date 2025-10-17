let
    // Try to get fresh data from API
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "daqbook-offsets/"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"point", "reading", "temp", "tn"}, {"point", "reading", "temp", "tn"}),
            #"Added Custom" = Table.AddColumn(#"Expanded Column1", "Offset", each [temp]-[reading]),
            #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"temp", "Temp"}, {"tn", "traceability_no"}}),
            #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Offset", type number}, {"Temp", type number}, {"reading", type number}, {"traceability_no", type text}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try Excel.CurrentWorkbook(){[Name="AllDaqbookOffsets"]}[Content] otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"point", "reading", "Temp", "traceability_no", "Offset"},
                    {{0, 0.0, 0.0, "Data unavailable: Unable to retrieve data from API or cached workbook. Please check your network connection and try again.", 0.0}}
                )
in
    FinalData