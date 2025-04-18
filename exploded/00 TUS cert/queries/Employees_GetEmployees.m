let
    Model_searchString = try Text.From(Excel.CurrentWorkbook(){[Name="Model.searchString"]}[Content][Column1]{0}) otherwise "",
    model.searchString = if Text.Length(Model_searchString) > 0 then [ model.searchString = Model_searchString ] else [],
    QueryOptions = Record.Combine({model.searchString}),
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/employees",
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