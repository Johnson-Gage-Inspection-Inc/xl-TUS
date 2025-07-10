let
    WorkItemNumber = try Text.From(Excel.CurrentWorkbook(){[Name="WorkItemNumber"]}[Content][Column1]{0}) otherwise "",
    
    // Early exit if no WorkItemNumber - return empty table with expected structure
    Result = if Text.Length(WorkItemNumber) = 0 then
        #table({"Name", "Value"}, {})
    else
        let
            // Protected API call - only include Query parameter when we have a WorkItemNumber
            Source = try Json.Document(
                Web.Contents(
                    "https://jgiapi.com",
                    [
                        RelativePath = "work-item-details",
                        Query = [ workItemNumber = WorkItemNumber ]
                    ]
                )
            ) otherwise null,
            
            // Handle API failure or null response
            ProcessedData = if Source = null then
                #table({"Name", "Value"}, {})
            else
                let
                    #"Converted to Table" = try Record.ToTable(Source) otherwise #table({"Name", "Value"}, {}),
                    #"Filtered Rows" = if Table.RowCount(#"Converted to Table") = 0 then 
                        #table({"Name", "Value"}, {})
                    else
                        Table.SelectRows(#"Converted to Table", each ([Name] <> "assetAttributes")),
                    
                    // Protected access to first row value
                    Value = if Table.RowCount(#"Converted to Table") > 0 then
                        try #"Converted to Table"{0}[Value] otherwise null
                    else
                        null,
                    
                    #"Converted to Table1" = if Value = null then
                        #table({"Name", "Value"}, {})
                    else
                        try Record.ToTable(Value) otherwise #table({"Name", "Value"}, {}),
                    
                    #"Appended Query" = try Table.Combine({#"Filtered Rows", #"Converted to Table1"}) otherwise #table({"Name", "Value"}, {})
                in
                    #"Appended Query"
        in
            ProcessedData
in
    Result