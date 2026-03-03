// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     True
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True

let
    // Try to get fresh data from API
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "wire-offsets/"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column2" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"correction_factor", "nominal_temp", "traceability_no", "wire_type"}, {"correction_factor", "nominal_temp", "traceability_no", "wire_type"}),
            #"Renamed Columns" = Table.RenameColumns(#"Expanded Column2",{{"correction_factor", "Offset"}, {"nominal_temp", "Temp"}}),
            #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Offset", type number}, {"Temp", type number}, {"traceability_no", type text}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try 
        let
            cached = Excel.CurrentWorkbook(){[Name="AllWireOffsets"]}[Content],
            selected = Table.SelectColumns(cached, {"traceability_no", "Temp", "Offset", "wire_type"}, MissingField.Error),
            validated = if Table.RowCount(selected) > 0 then selected else null
        in validated
    otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"traceability_no", "Temp", "Offset", "wire_type"},
                    {}
                ),
    #"Reordered Columns" = Table.ReorderColumns(FinalData,{"traceability_no", "Temp", "Offset", "wire_type"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"traceability_no", Order.Ascending}, {"Temp", Order.Ascending}})
in
    #"Sorted Rows"