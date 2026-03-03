let
    schema = {"asset_tag", "asset_id", "certificate_tn", "due_date", "service_date"},

    // --- Try to get fresh data from API ---
    FreshData = try
        let
            response = Web.Contents(
                "https://jgiapi.com",
                [ RelativePath = "daqbook-service-records/latest/" ]
            ),
            json = Json.Document(response),
            #"Converted to Table" = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"asset_id", "certificate_tn", "due_date", "service_date"}, {"asset_id", "certificate_tn", "due_date", "service_date"}),
            #"Changed Type" = Table.TransformColumnTypes(#"Expanded Column1",{{"service_date", type datetime}, {"due_date", type datetime}, {"asset_id", Int64.Type}, {"certificate_tn", type text}}),
            #"Inserted Text Before Delimiter" = Table.AddColumn(#"Changed Type", "asset_tag", each try Text.BeforeDelimiter([certificate_tn], "_") otherwise [certificate_tn], type text),
            #"Reordered Columns" = Table.ReorderColumns(#"Inserted Text Before Delimiter", schema),
            #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns",{{"service_date", type date}, {"due_date", type date}})
        in
            #"Changed Type1"
    otherwise null,

    // --- Fallback: cached data from workbook table ---
    CachedData = try
        let
            cached    = Excel.CurrentWorkbook(){[Name="DaqbookServiceRecords"]}[Content],
            selected  = Table.SelectColumns(cached, schema),
            validated = if Table.RowCount(selected) > 0 then selected else null
        in
            validated
    otherwise null,

    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData
                else if CachedData <> null then CachedData
                else #table(schema, {})
in
    FinalData