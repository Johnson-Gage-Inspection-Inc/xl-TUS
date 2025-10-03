let
    // Try to get fresh data from SharePoint
    FreshData = try 
        let
            Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/ClientOvens.xlsx"), null, true),
            Header_Info_Table = Source{[Item="UUTRangeTol",Kind="Table"]}[Data],
            #"Changed Type" = Table.TransformColumnTypes(Header_Info_Table,{{"AssetId", Int64.Type}, {"RangeMin", Int64.Type}, {"RangeMax", Int64.Type}, {"TolMin", type number}, {"TolMax", type number}, {"TolResolution", type number}, {"Unit", type text}, {"Flags", type text}, {"Link (Qualer)", Int64.Type}}),
            #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Link (Qualer)", "Symmetrical Tolerance"})
        in
            #"Removed Columns"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try 
        let
            Source = Excel.CurrentWorkbook(){[Name="UUTRangeTol"]}[Content],
            #"Changed Type" = Table.TransformColumnTypes(Source,{{"AssetId", Int64.Type}, {"RangeMin", Int64.Type}, {"RangeMax", Int64.Type}, {"TolMin", type number}, {"TolMax", type number}, {"TolResolution", type number}, {"Unit", type text}, {"Flags", type text}, {"TolMinText", type text}, {"TolMaxText", type text}, {"TolText", type text}, {"RangeMaxText", type text}, {"RangeText", type text}}),
            #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"Flags", "Unit", "TolResolution", "TolMax", "TolMin", "RangeMax", "RangeMin", "AssetId"})
        in
            #"Removed Other Columns"
    otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return fallback table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"AssetId", "RangeMin", "RangeMax", "TolMin", "TolMax", "TolResolution", "Unit", "Flags"},
                    {{0, 0, 0, 0.0, 0.0, 0.0, "Data unavailable: Unable to retrieve data from SharePoint or cached workbook. Please check your network connection and try again.", ""}}
                )
in
    FinalData