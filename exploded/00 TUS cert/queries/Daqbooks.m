let
    Source = AssetPool,
    #"Filtered Rows" = Table.SelectRows(Source, each ([RootCategoryName] = "DAQ Modules")),
    #"Sorted Rows" = Table.Sort(#"Filtered Rows",{{"AssetTag", Order.Ascending}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Sorted Rows",{"AssetId", "SerialNumber", "AssetTag", "AssetDescription", "Notes", "ManufacturerPartNumber"})
in
    #"Removed Other Columns"