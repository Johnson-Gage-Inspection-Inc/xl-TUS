// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     True
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    True

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
    CachedData = try 
        let
            cached = Excel.CurrentWorkbook(){[Name="AllDaqbookOffsets"]}[Content],
            validated = if Table.RowCount(cached) > 0 then cached else null
        in validated
    otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"point", "reading", "Temp", "traceability_no", "Offset"},
                    {}
                ),
    #"Reordered Columns" = Table.ReorderColumns(FinalData,{"traceability_no", "point", "Temp", "reading", "Offset"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"traceability_no", Order.Ascending}, {"Temp", Order.Ascending}, {"point", Order.Ascending}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Sorted Rows",{"traceability_no", "reading", "point", "Temp", "Offset"})
in
    #"Reordered Columns1"