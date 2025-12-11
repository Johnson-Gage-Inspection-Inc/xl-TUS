let
    // Try to get fresh data from API
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "pyro-assets"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column2" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"ActivationDate", "AssetDescription", "AssetId", "AssetMaker", "AssetName", "AssetStatus", "AssetTag", "AssetUser", "Barcode", "CategoryName", "Class", "CompanyId", "Condition", "Criticality", "CumulativeServiceCost", "CustodianEmail", "CustodianName", "CustomStatusAbbr", "CustomStatusName", "Department", "DepreciationBasis", "DepreciationMethod", "DocumentNumber", "DocumentSection", "EquipmentId", "LegacyIdentifier", "LifeSpanMonths", "Location", "ManufacturerPartNumber", "Notes", "Pool", "ProductDescription", "ProductId", "ProductManufacturer", "ProductName", "PurchaseCost", "PurchaseDate", "RetirementDate", "RetirmentReason", "RoomNumber", "RootCategoryName", "SalvageValue", "SerialNumber", "SiteId", "SiteName", "Station"}, {"ActivationDate", "AssetDescription", "AssetId", "AssetMaker", "AssetName", "AssetStatus", "AssetTag", "AssetUser", "Barcode", "CategoryName", "Class", "CompanyId", "Condition", "Criticality", "CumulativeServiceCost", "CustodianEmail", "CustodianName", "CustomStatusAbbr", "CustomStatusName", "Department", "DepreciationBasis", "DepreciationMethod", "DocumentNumber", "DocumentSection", "EquipmentId", "LegacyIdentifier", "LifeSpanMonths", "Location", "ManufacturerPartNumber", "Notes", "Pool", "ProductDescription", "ProductId", "ProductManufacturer", "ProductName", "PurchaseCost", "PurchaseDate", "RetirementDate", "RetirmentReason", "RoomNumber", "RootCategoryName", "SalvageValue", "SerialNumber", "SiteId", "SiteName", "Station"}),
            #"Filtered Rows1" = Table.SelectRows(#"Expanded Column2", each ([AssetStatus] = "Active")),
            #"Filtered Rows" = Table.SelectRows(#"Filtered Rows1", each ([RetirmentReason] = null)),
            #"Removed Columns1" = Table.RemoveColumns(
            #"Filtered Rows",
            {
              "RetirementDate",
              "RetirmentReason",
              "CompanyId",
              "EquipmentId",
              "AssetStatus",
              "Criticality",
              "Pool",
              "PurchaseDate",
              "PurchaseCost",
              "LifeSpanMonths",
              "ActivationDate",
              "DepreciationBasis",
              "DepreciationMethod",
              "SalvageValue",
              "Condition",
              "SiteName",
              "SiteId",
              "DocumentNumber",
              "DocumentSection",
              "CumulativeServiceCost",
              "ProductId",
              "Station",
              "CustodianEmail",
              "Barcode",
              "LegacyIdentifier",
              "CustodianName",
              "Location",
              "RoomNumber"
            }
        ),
            #"Sorted Rows" = Table.Sort(#"Removed Columns1",{{"AssetTag", Order.Ascending}})
        in
            #"Sorted Rows"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try Excel.CurrentWorkbook(){[Name="AssetPool"]}[Content] otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"AssetDescription", "AssetId", "AssetMaker", "AssetName", "AssetTag", "AssetUser", "CategoryName", "CustodianFirstName", "CustodianLastName", "Department", "ManufacturerPartNumber", "Notes", "ProductDescription", "ProductManufacturer", "ProductName", "RootCategoryName", "SerialNumber"},
                    {{"No data available", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}}
                ),
    #"Filtered Rows" = Table.SelectRows(FinalData, each ([SerialNumber] <> ""))
in
    #"Filtered Rows"