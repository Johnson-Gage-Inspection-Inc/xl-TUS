let
    AssetId = Text.From(Excel.CurrentWorkbook(){[Name="AssetId"]}[Content][Column1]{0}),
    QueryOptions = [],
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/clients/assets/" & AssetId & "/attributes",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Query = QueryOptions,
            Headers = [ Authorization = "Api-Token REDACTED" ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json)
in
    ConvertToTable