let
    // Define the required column structure
    requiredColumns = {"employee_id", "Name", "login_email", "IsPyro"},
    emptyTable = #table(requiredColumns, {}),
    
    // First, try to get existing cached employees data
    Source = try Employees otherwise Excel.CurrentWorkbook(){[Name="Employees"]}[Content],
    
    // First, try to get existing cached employee data
    existingEmployeeData = try Excel.CurrentWorkbook(){[Name="Employee"]}[Content] otherwise emptyTable,
    
    // Try to get fresh employee data from API, fall back to cached data if it fails
    employeeData = try 
        let
            response = Web.Contents("https://jgiapi.com", [
                RelativePath = "whoami",
                ManualStatusHandling = {400, 401, 403, 404, 500, 502, 503, 504}
            ]),
            json = Json.Document(response),
            email = json[user],
            // Filter the employees data to get the current user
            #"Filtered Rows" = try Table.SelectRows(Source, each try [login_email] = email otherwise false) otherwise Source
        in
            #"Filtered Rows"
    otherwise
        // If API call fails, use cached employee data
        if Table.RowCount(existingEmployeeData) > 0 then existingEmployeeData else emptyTable,
    
    // Ensure the result always has the required columns with proper structure
    finalResult = if Table.RowCount(employeeData) > 0 then
        let
            // Get existing columns from the data
            existingColumns = Table.ColumnNames(employeeData),
            // Add missing columns with null values
            tableWithAllColumns = List.Accumulate(
                requiredColumns,
                employeeData,
                (table, column) => 
                    if List.Contains(existingColumns, column) then 
                        table 
                    else 
                        Table.AddColumn(table, column, each null)
            ),
            // Select only the required columns in the correct order
            finalTable = Table.SelectColumns(tableWithAllColumns, requiredColumns)
        in
            finalTable
    else
        emptyTable
in
    finalResult