let
    Source = Employees_GetEmployees,
    #"Expanded Departments" = Table.ExpandListColumn(Source, "Departments"),
    #"Expanded Departments1" = Table.ExpandRecordColumn(#"Expanded Departments", "Departments", {"Id", "Name", "Position"}, {"Departments.Id", "Departments.Name", "Departments.Position"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Departments1", each ([Departments.Id] = 2278) and ([IsDeleted] = false)),
    #"Merged Columns" = Table.CombineColumns(#"Filtered Rows",{"FirstName", "LastName"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Name"),
    #"Removed Other Columns" = Table.SelectColumns(#"Merged Columns",{"Name", "EmployeeId"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"EmployeeId", "Name"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns",{{"Name", Order.Ascending}})
in
    #"Sorted Rows"