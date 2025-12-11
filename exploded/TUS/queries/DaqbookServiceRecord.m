let
    // Try to get the Asset ID
    assetIdCell = try Excel.CurrentWorkbook(){[Name="DaqbookAssetId"]}[Content][Column1]{0} otherwise null,
    DaqbookAssetId = if assetIdCell = null or Text.Trim(Text.From(assetIdCell)) = "" then null else Text.From(assetIdCell),

    // First, try to get existing cached data from the table
    existingDataRaw = try Excel.CurrentWorkbook(){[Name="DaqbookServiceRecord"]}[Content] otherwise null,
    existingData = if existingDataRaw <> null and Table.ColumnCount(existingDataRaw) = 2 then
        Table.RenameColumns(existingDataRaw, {
            {Table.ColumnNames(existingDataRaw){0}, "Name"},
            {Table.ColumnNames(existingDataRaw){1}, "Value"}
        })
    else
        #table({"Name", "Value"}, {}),
    #"Filtered Rows1" = Table.SelectRows(existingData, each ([Name] <> null)),
    
    // Handle missing ID by returning cached data or empty table
    result = if DaqbookAssetId = null then
        // Return cached data if available, otherwise empty table
        if Table.RowCount(#"Filtered Rows1") > 1 then #"Filtered Rows1" else #table({"Name", "Value"}, {})
    else
        // Try to fetch fresh data, fall back to cached data if fetch fails
        let
            freshData = try 
                let
                    relativeUrl = "asset-service-records/" & DaqbookAssetId,
                    // Use ManualStatusHandling to prevent auto-refresh failures
                    response = Web.Contents("https://jgiapi.com", [
                        RelativePath = relativeUrl,
                        ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503, 504}
                    ]),
                    json = Json.Document(response),
                    ConvertToTable = if Value.Is(json, type list) then Table.FromRecords(json) else Record.ToTable(json),
                    #"Changed Type1" = Table.TransformColumnTypes(ConvertToTable,{{"service_date", type datetime}}),
                    #"Sorted Rows" = Table.Sort(#"Changed Type1",{{"service_date", Order.Descending}}),
                    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"service_date", type datetime}, {"due_date", type datetime}, {"next_service_date", type datetime}}),
                    #"Filtered Rows" = Table.TransformColumns(#"Changed Type", {
                        {"service_date", DateTime.Date},
                        {"due_date", DateTime.Date},
                        {"next_service_date", DateTime.Date}
                    }){0},
                    #"Converted to Table" = Record.ToTable(#"Filtered Rows")
                in
                    #"Converted to Table"
            otherwise 
                // If web call fails, use cached data or empty table
                if Table.RowCount(#"Filtered Rows1") > 1 then #"Filtered Rows1" else #table({"Name", "Value"}, {})
        in
            freshData,
    #"Filtered Rows" = Table.SelectRows(result, each ([Name] <> "additional_properties"))
in
    #"Filtered Rows"