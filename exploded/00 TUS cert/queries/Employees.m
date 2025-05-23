let
    Model_searchString = try Text.From(Excel.CurrentWorkbook(){[Name="Model.searchString"]}[Content][Column1]{0}) otherwise "",
    model.searchString = if Text.Length(Model_searchString) > 0 then [ model.searchString = Model_searchString ] else [],
    QueryOptions = Record.Combine({model.searchString}),
    baseUrl = "https://jgiquality.qualer.com",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = "api/employees",
            Query = QueryOptions,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json),
    #"Expanded Departments" = Table.ExpandListColumn(ConvertToTable, "Departments"),
    #"Expanded Departments1" = Table.ExpandRecordColumn(#"Expanded Departments", "Departments", {"Id", "Name", "Position"}, {"Departments.Id", "Departments.Name", "Departments.Position"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Departments1", each ([Departments.Id] = 2278) and ([IsDeleted] = false)),
    #"Merged Columns" = Table.CombineColumns(#"Filtered Rows",{"FirstName", "LastName"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Name"),
    #"Removed Other Columns" = Table.SelectColumns(#"Merged Columns",{"Name", "EmployeeId"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"EmployeeId", "Name"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Name", Order.Ascending}})
in
    #"Sorted Rows"