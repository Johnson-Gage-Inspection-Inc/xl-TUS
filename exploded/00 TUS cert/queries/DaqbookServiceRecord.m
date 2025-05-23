let
    DaqbookAssetId = Text.From(Excel.CurrentWorkbook(){[Name="DaqbookAssetId"]}[Content][Column1]{0}),
    QueryOptions = [],
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/assets/" & DaqbookAssetId & "/assetservicerecords",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Query = QueryOptions,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = if Value.Is(json, type list) then Table.FromRecords(json) else Record.ToTable(json),
    #"Sorted Rows" = Table.Sort(ConvertToTable,{{"ServiceDate", Order.Descending}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Sorted Rows",{{"ServiceDate", type datetime}, {"DueDate", type datetime}, {"NextServiceDate", type datetime}}),
    #"Filtered Rows" = Table.TransformColumns(#"Changed Type", {
        {"ServiceDate", DateTime.Date},
        {"DueDate", DateTime.Date},
        {"NextServiceDate", DateTime.Date}
    }){0},
    #"Converted to Table" = Record.ToTable(#"Filtered Rows")
in
    #"Converted to Table"