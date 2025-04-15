let
    ClientCompanyId = Excel.CurrentWorkbook(){[Name="ClientCompanyId"]}[Content]{0}[Column1],
    relativeUrl = "api/service/clients/" & Text.From(ClientCompanyId) & "/assets",
    response = Web.Contents(
        "https://jgiquality.qualer.com",
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Headers = [ Authorization = "Api-Token bf407589-f463-4046-ba2c-30642bd5d637" ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json),
    FillRootCategories = Table.AddColumn(
        ConvertToTable,
        "RootCategoryNameFilled",
        each if Value.Is([RootCategoryName], type text) then [RootCategoryName]
            else if Value.Is([CategoryName], type text) then [CategoryName]
            else null
    ),
    #"Filtered Rows" = Table.SelectRows(FillRootCategories, each (
        [RootCategoryNameFilled] = "Furnaces" or
        // [RootCategoryNameFilled] = "Hardware Tools" or
        // [RootCategoryNameFilled] = "Heating Equipment" or
        // [RootCategoryNameFilled] = "Reagents, Chemicals, & Gases" or
        // [RootCategoryNameFilled] = "Sealers" or
        // [RootCategoryNameFilled] = "Spray Dryer" or
        [RootCategoryNameFilled] = "Thermometers")
        )
in
    #"Filtered Rows"