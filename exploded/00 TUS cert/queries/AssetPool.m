let
    response = Web.Contents(
        "https://jgiquality.qualer.com",
        [
            RelativePath = "api/assets/byassetpool/620646",
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json),
    #"Removed Columns" = Table.RemoveColumns(ConvertToTable,{"CompositeParentId", "CompositeChildCount"}),
    #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([AssetStatus] = 1)),
    #"Filtered Rows" = Table.SelectRows(#"Filtered Rows1", each ([RetirmentReason] = null)),
    #"Removed Columns1" = Table.RemoveColumns(#"Filtered Rows",{"RetirementDate", "RetirmentReason", "CompanyId", "EquipmentId", "AssetStatus", "Criticality", "Pool", "PurchaseDate", "PurchaseCost", "LifeSpanMonths", "ActivationDate", "DepreciationBasis", "DepreciationMethod", "SalvageValue", "Condition", "SiteName", "SiteId", "DocumentNumber", "DocumentSection", "CumulativeServiceCost", "ProductId", "Station", "Class", "CustodianEmail", "Barcode", "LegacyIdentifier", "CustodianName", "Location", "RoomNumber"})
in
    #"Removed Columns1"