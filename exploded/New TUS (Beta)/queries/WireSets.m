let
    Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/WireSetCerts.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Removed Other Columns" = Table.SelectColumns(#"Promoted Headers",{"type", "wire_roll_cert_number", "certificate_number", "next_service_date", "service_date", "custom_order_number", "asset_tag", "serial_number", "asset_id"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Removed Other Columns",{{"next_service_date", type date}, {"service_date", type date}, {"asset_id", Int64.Type}})
in
    #"Changed Type"