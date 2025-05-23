let
    QueryOptions = [],
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/workorders/" & Text.From(#"ServiceOrderId") & "",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Query = QueryOptions,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = if Value.Is(json, type list) then Table.FromRecords(json) else Record.ToTable(json)
in
    ConvertToTable