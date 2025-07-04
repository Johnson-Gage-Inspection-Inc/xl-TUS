let
    // Try to get the Asset ID
    assetIdCell = try Excel.CurrentWorkbook(){[Name="DaqbookAssetId"]}[Content][Column1]{0} otherwise null,
    DaqbookAssetId = if assetIdCell = null or Text.Trim(Text.From(assetIdCell)) = "" then null else Text.From(assetIdCell),

    // Handle missing ID by returning empty table early
    result = if DaqbookAssetId = null then
        #table({"Name", "Value"}, {})  // empty table with expected structure
    else
        let
            relativeUrl = "asset-service-records/" & DaqbookAssetId,
            response = Web.Contents("https://api.jgiquality.com", [RelativePath = relativeUrl]),
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
in
    result