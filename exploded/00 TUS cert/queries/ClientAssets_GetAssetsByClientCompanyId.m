let
    ClientCompanyId = Excel.CurrentWorkbook(){[Name="ClientCompanyId"]}[Content]{0}[Column1],
    relativeUrl = "api/service/clients/" & Text.From(ClientCompanyId) & "/assets",
    response = Web.Contents(
        "https://jgiquality.qualer.com",
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Headers = [ Authorization = "Api-Token REDACTED" ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json),
    #"Filtered Rows" = Table.SelectRows(ConvertToTable, each [RootCategoryName] = "Furnaces" or [CategoryName] = "Furnaces")
in
    #"Filtered Rows"