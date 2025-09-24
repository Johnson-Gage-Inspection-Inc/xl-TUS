let
    // Try to get fresh data from SharePoint
    FreshData = try 
        let
            Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/WireSetCerts.xlsx"), null, true),
            Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
            #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
            #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"service_date", type date}, {"asset_id", Int64.Type}})
        in
            #"Changed Type"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try
        let
            existing = Excel.CurrentWorkbook(){[Name="WireSets"]}[Content],
            #"Changed Type" = Table.TransformColumnTypes(existing,{{"service_date", type date}})
        in
            #"Changed Type"
        otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"asset_id",  "serial_number", "asset_tag", "custom_order_number", "service_date", "certificate_number", "wire_roll_cert_number", "type"},
                    {{"No data available", "", "", null, null, "", "", "", 0}}
                )
in
    FinalData