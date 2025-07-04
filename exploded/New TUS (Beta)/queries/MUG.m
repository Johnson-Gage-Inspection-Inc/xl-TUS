let
    Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/General/MUG Tables.xlsm"), null, true),
    Thermodynamic_Table = Source{[Item="Thermodynamic",Kind="Table"]}[Data],
    #"Changed Type" = Table.TransformColumnTypes(Thermodynamic_Table,{{"Item", type text}, {"Range Lo", Int64.Type}, {"Range Hi", Int64.Type}, {"Res", type number}, {"Meas Units", type text}, {"Uncert Units", type text}, {"Base ", type number}, {"Mult", Int64.Type}, {"Base 2", type number}, {"Mult 2", Int64.Type}, {"Repeat Mult", Int64.Type}, {"Item_C", type text}, {"File", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each ([Item] = "TUS")),
    #"Removed Other Columns" = Table.SelectColumns(#"Filtered Rows",{"Range Lo", "Range Hi", "Res", "Meas Units", "Uncert Units", "Base ", "Mult", "Base 2", "Mult 2"})
in
    #"Removed Other Columns"