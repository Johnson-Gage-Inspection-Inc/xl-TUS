let
    // Try to get fresh data from SharePoint
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "wire-set-certs/"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"asset_id", "serial_number", "asset_tag", "custom_order_number", "service_date", "certificate_number", "traceability_number", "wire_roll_cert_number", "wire_set_group"}, {"asset_id", "serial_number", "asset_tag", "custom_order_number", "service_date", "certificate_number", "traceability_number", "wire_roll_cert_number", "type"}),
            #"Sorted Rows" = Table.Sort(#"Expanded Column1",{{"asset_tag", Order.Ascending}}),
            #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"service_date", type datetime}, {"asset_id", Int64.Type}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try
        let
            existing = Excel.CurrentWorkbook(){[Name="WireSets"]}[Content],
            #"Changed Type" = Table.TransformColumnTypes(existing,{{"service_date", type date}})
        in
            #"Changed Type"
        otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"asset_id",  "serial_number", "asset_tag", "custom_order_number", "service_date", "certificate_number", "traceability_number", "wire_roll_cert_number", "type"},
                    {{0, "No data available", "", "", null, "", "", "", ""}}
                )
in
    FinalData