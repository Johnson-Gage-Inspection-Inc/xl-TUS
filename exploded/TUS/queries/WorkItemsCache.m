// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    False

let
    WorkOrderNumber = try Text.From(Excel.CurrentWorkbook(){[Name="WorkOrderNumber"]}[Content][Column1]{0}) otherwise "",

    CacheColumns = {"AsFoundCheck", "AsFoundResult", "AsLeftCheck", "AsLeftResult", "AssetCompanyId", "AssetCompanyName", "AssetDescription", "AssetId", "AssetName", "AssetOwnership", "AssetSiteId", "AssetSiteName", "AssetTag", "CertificateNumber", "ChannelCount", "CheckedById", "CheckedByName", "CheckedOn", "ClientCompanyId", "CompletedById", "CompletedByName", "CompletedOn", "CreatedOnUtc", "CustomOrderNumber", "CustomWorkStatus", "DocumentNumber", "DocumentSection", "DueDate", "IsLimited", "ItemAsFoundResult", "ItemAsLeftResult", "ItemResultStatus", "MaintenancePlan", "MaintenanceTask", "NextServiceDate", "NextServiceLevel", "OrderItemNumber", "OverridePartsTotal", "OverrideRepairsTotal", "OverrideServiceTotal", "PartsTotal", "PartsTotalBeforeDiscount", "ProviderCompany", "ProviderPhone", "ProviderTechnician", "RepairsTotal", "ResultStatus", "SerialNumber", "ServiceCharge", "ServiceComments", "ServiceDate", "ServiceLevel", "ServiceLevelDocumentNumber", "ServiceNotes", "ServiceOptions", "ServiceOrderId", "ServiceTotal", "ServiceType", "UpdatedBy", "UpdatedById", "UpdatedOnUtc", "VendorCompanyId", "WorkItemId", "WorkStatus"},

    EmptyCache = #table(CacheColumns, {}),

    FreshData =
        if Text.Length(WorkOrderNumber) = 0 then
            null
        else
            try
                let
                    Source = Json.Document(
                        Web.Contents(
                            "https://jgiapi.com",
                            [
                                RelativePath = "api/service/workitems",
                                Query = [workItemNumber = WorkOrderNumber]
                            ]
                        )
                    )
                in
                    if Value.Is(Source, type list) then
                        try Table.FromRecords(Source, CacheColumns, MissingField.UseNull) otherwise null
                    else
                        null
            otherwise
                null,

    CachedData =
        try
            let
                cached = Excel.CurrentWorkbook(){[Name="WorkItemsCache"]}[Content],
                validated = if Table.RowCount(cached) > 0 then Table.SelectColumns(cached, CacheColumns, MissingField.UseNull) else null
            in
                validated
        otherwise
            null,

    FinalData =
        if Text.Length(WorkOrderNumber) = 0 then EmptyCache
        else if FreshData <> null then FreshData
        else if CachedData <> null then CachedData
        else EmptyCache
in
    FinalData