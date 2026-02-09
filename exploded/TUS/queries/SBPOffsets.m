let
    // Try to get fresh data from API
    FreshData = try 
        let
            response = Json.Document(Web.Contents("https://jgiapi.com/sbp-offsets/")),
            #"Converted to Table" = Table.FromList(response, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"certificate_number", "nominal_temp", "offset", "probe_serial_number"}, {"certificate_number", "nominal_temp", "offset", "probe_serial_number"}),
            #"Changed Type" = Table.TransformColumnTypes(#"Expanded Column1",{{"probe_serial_number", type text}, {"certificate_number", type text}, {"nominal_temp", type number}, {"offset", type number}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try 
        let
            cached = Excel.CurrentWorkbook(){[Name="SBPOffsets"]}[Content],
            #"Removed Other Columns" = Table.SelectColumns(cached,{"certificate_number", "nominal_temp", "offset", "probe_serial_number"})
        in
            #"Removed Other Columns"
    otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"probe_serial_number", "certificate_number", "nominal_temp", "offset"},
                    {{"Data unavailable: Unable to retrieve data from API or cached workbook. Please check your network connection and try again.", "", 0.0, 0.0}}
                ),
    #"Sorted Rows" = Table.Sort(FinalData,{{"certificate_number", Order.Ascending}, {"probe_serial_number", Order.Ascending}, {"nominal_temp", Order.Ascending}})
in
    #"Sorted Rows"