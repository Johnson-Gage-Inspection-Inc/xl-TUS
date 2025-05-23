let
    QueryOptions = [],
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/clients/assets/" & Text.From(AssetId) & "/attributes",
    response = try Web.Contents(
        baseUrl,
        [
            RelativePath = relativeUrl,
            Query = QueryOptions,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = try Json.Document(response),
    ConvertToTable = try Table.FromRecords(json),
    Output = if ConvertToTable[HasError] or not Value.Is(ConvertToTable[Value], type table)
             then #table({"Name", "Value"}, {{"", ""}})
             else ConvertToTable[Value]
in
    Output