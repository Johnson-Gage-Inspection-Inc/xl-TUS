let
    assetId = if Text.Length(Text.From(AssetId)) > 0 then Text.From(AssetId) else "",
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/clients/assets/" & assetId,
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = relativeUrl,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    #"Converted to Table" = Record.ToTable(json)
in
    #"Converted to Table"