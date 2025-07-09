let
    // Try to get fresh data from API
    FreshData = try 
        let
            Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "pyro-assets"])),
            #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
            #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"activation_date", "asset_description", "asset_id", "asset_maker", "asset_name", "asset_status", "asset_tag", "asset_user", "barcode", "category_name", "company_id", "composite_child_count", "composite_parent_id", "condition", "criticality", "cumulative_service_cost", "custodian_email", "custodian_first_name", "custodian_last_name", "custodian_name", "department", "depreciation_basis", "depreciation_method", "document_number", "document_section", "equipment_id", "legacy_identifier", "life_span_months", "location", "manufacturer_part_number", "notes", "pool", "product_description", "product_id", "product_manufacturer", "product_name", "purchase_cost", "purchase_date", "retirement_date", "retirment_reason", "room_number", "root_category_name", "salvage_value", "serial_number", "site_id", "site_name", "station"}, {"activation_date", "asset_description", "asset_id", "asset_maker", "asset_name", "asset_status", "asset_tag", "asset_user", "barcode", "category_name", "company_id", "composite_child_count", "composite_parent_id", "condition", "criticality", "cumulative_service_cost", "custodian_email", "custodian_first_name", "custodian_last_name", "custodian_name", "department", "depreciation_basis", "depreciation_method", "document_number", "document_section", "equipment_id", "legacy_identifier", "life_span_months", "location", "manufacturer_part_number", "notes", "pool", "product_description", "product_id", "product_manufacturer", "product_name", "purchase_cost", "purchase_date", "retirement_date", "retirment_reason", "room_number", "root_category_name", "salvage_value", "serial_number", "site_id", "site_name", "station"}),
            #"Removed Columns" = Table.RemoveColumns(#"Expanded Column1",{"composite_parent_id", "composite_child_count"}),
            #"Filtered Rows1" = Table.SelectRows(#"Removed Columns", each ([asset_status] = "Active")),
            #"Filtered Rows" = Table.SelectRows(#"Filtered Rows1", each ([retirment_reason] = null)),
            #"Removed Columns1" = Table.RemoveColumns(
            #"Filtered Rows",
            {
              "retirement_date",
              "retirment_reason",
              "company_id",
              "equipment_id",
              "asset_status",
              "criticality",
              "pool",
              "purchase_date",
              "purchase_cost",
              "life_span_months",
              "activation_date",
              "depreciation_basis",
              "depreciation_method",
              "salvage_value",
              "condition",
              "site_name",
              "site_id",
              "document_number",
              "document_section",
              "cumulative_service_cost",
              "product_id",
              "station",
              "custodian_email",
              "barcode",
              "legacy_identifier",
              "custodian_name",
              "location",
              "room_number"
            }
        ),
            #"Sorted Rows" = Table.Sort(#"Removed Columns1",{{"asset_tag", Order.Ascending}})
        in
            #"Sorted Rows"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try Excel.CurrentWorkbook(){[Name="AssetPool"]}[Content] otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"asset_description", "asset_id", "asset_maker", "asset_name", "asset_tag", "asset_user", "category_name", "custodian_first_name", "custodian_last_name", "department", "manufacturer_part_number", "notes", "product_description", "product_manufacturer", "product_name", "root_category_name", "serial_number"},
                    {{"No data available", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}}
                ),
    #"Filtered Rows" = Table.SelectRows(FinalData, each ([serial_number] <> ""))
in
    #"Filtered Rows"