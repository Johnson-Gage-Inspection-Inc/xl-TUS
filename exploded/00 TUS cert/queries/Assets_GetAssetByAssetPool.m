let
    AssetPoolId = 620646,
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/assets/byassetpool/" & Text.From(AssetPoolId) & "",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Headers = [ Authorization = "Api-Token REDACTED" ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column2" = Table.ExpandRecordColumn(ConvertToTable, "Column1", {"CompanyId", "AssetId", "SerialNumber", "AssetUser", "AssetTag", "EquipmentId", "AssetStatus", "AssetName", "AssetDescription", "AssetMaker", "Location", "RoomNumber", "Barcode", "LegacyIdentifier", "RootCategoryName", "CategoryName", "Class", "CustodianEmail", "CustodianName", "Department", "Station", "Notes", "DocumentNumber", "DocumentSection", "CumulativeServiceCost", "ProductId", "ManufacturerPartNumber", "ProductName", "ProductDescription", "ProductManufacturer", "SiteName", "SiteId", "Condition", "Criticality", "Pool", "PurchaseDate", "PurchaseCost", "LifeSpanMonths", "ActivationDate", "DepreciationBasis", "DepreciationMethod", "RetirementDate", "SalvageValue", "RetirmentReason", "CompositeParentId", "CompositeChildCount"}, {"CompanyId", "AssetId", "SerialNumber", "AssetUser", "AssetTag", "EquipmentId", "AssetStatus", "AssetName", "AssetDescription", "AssetMaker", "Location", "RoomNumber", "Barcode", "LegacyIdentifier", "RootCategoryName", "CategoryName", "Class", "CustodianEmail", "CustodianName", "Department", "Station", "Notes", "DocumentNumber", "DocumentSection", "CumulativeServiceCost", "ProductId", "ManufacturerPartNumber", "ProductName", "ProductDescription", "ProductManufacturer", "SiteName", "SiteId", "Condition", "Criticality", "Pool", "PurchaseDate", "PurchaseCost", "LifeSpanMonths", "ActivationDate", "DepreciationBasis", "DepreciationMethod", "RetirementDate", "SalvageValue", "RetirmentReason", "CompositeParentId", "CompositeChildCount"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Column2",{"CompositeParentId", "CompositeChildCount"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([AssetStatus] = 1)),
    #"Filtered Rows" = Table.SelectRows(#"Filtered Rows1", each ([RetirmentReason] = null)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows",{"RetirementDate", "RetirmentReason", "CompanyId", "EquipmentId", "AssetStatus", "Criticality", "Pool", "PurchaseDate", "PurchaseCost", "LifeSpanMonths", "ActivationDate", "DepreciationBasis", "DepreciationMethod", "SalvageValue", "Condition", "SiteName", "SiteId", "DocumentNumber", "DocumentSection", "CumulativeServiceCost", "ProductId", "Station", "Class", "CustodianEmail", "Barcode", "LegacyIdentifier", "CustodianName", "Location", "RoomNumber"})
in
    #"Removed Columns1"