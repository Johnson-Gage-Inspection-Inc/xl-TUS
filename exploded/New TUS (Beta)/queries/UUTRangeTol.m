let
    Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/ClientOvens.xlsx"), null, true),
    Header_Info_Table = Source{[Item="UUTRangeTol",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Header_Info_Table,{{"AssetId", Int64.Type}, {"RangeMin", Int64.Type}, {"RangeMax", Int64.Type}, {"TolMin", type number}, {"TolMax", type number}, {"TolResolution", type number}, {"Unit", type text}, {"Flags", type text}, {"Link (Qualer)", Int64.Type}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Link (Qualer)"})
in
    #"Removed Columns"