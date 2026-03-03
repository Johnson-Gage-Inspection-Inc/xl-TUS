// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True

let
    // Try to get fresh data from SharePoint
    FreshData = try 
        let
            Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/General/MUG Tables.xlsm"), null, true),
            Thermodynamic_Table = Source{[Item="Thermodynamic",Kind="Table"]}[Data],
            #"Changed Type" = Table.TransformColumnTypes(Thermodynamic_Table,{{"Item", type text}, {"Range Lo", Int64.Type}, {"Range Hi", Int64.Type}, {"Res", type number}, {"Meas Units", type text}, {"Uncert Units", type text}, {"Base ", type number}, {"Mult", Int64.Type}, {"Base 2", type number}, {"Mult 2", Int64.Type}, {"Repeat Mult", Int64.Type}, {"Item_C", type text}, {"File", type text}}),
            #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Item] = "TUS")),
            #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Range Lo", "Range Hi", "Res", "Meas Units", "Uncert Units", "Base ", "Mult", "Base 2", "Mult 2"})
        in
            #"Removed Other Columns"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try 
        let
            cached = Excel.CurrentWorkbook(){[Name="MUG"]}[Content],
            validated = if Table.RowCount(cached) > 0 then cached else null
        in validated
    otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"Range Lo", "Range Hi", "Res", "Meas Units", "Uncert Units", "Base ", "Mult", "Base 2", "Mult 2"},
                    {}
                )
in
    FinalData