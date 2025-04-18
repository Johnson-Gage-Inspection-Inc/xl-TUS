let
    AssetPoolId = Text.From(Excel.CurrentWorkbook(){[Name="AssetPoolId"]}[Content]{0}[Column1]),
    QueryOptions = [],
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/assets/byassetpool/" & AssetPoolId & "",
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