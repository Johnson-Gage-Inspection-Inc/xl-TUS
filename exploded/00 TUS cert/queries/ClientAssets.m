let
    baseUrl = "https://jgiquality.qualer.com",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = "api/service/clients/assets/" & Text.From(AssetId),
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = if Value.Is(json, type list) then Table.FromRecords(json) else Record.ToTable(json)
in
    ConvertToTable