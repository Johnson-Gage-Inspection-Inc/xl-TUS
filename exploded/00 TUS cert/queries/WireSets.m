let
    Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/WireSetCerts.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"asset_id", Int64.Type}, {"serial_number", type text}, {"asset_tag", type text}, {"custom_order_number", type text}, {"service_date", type datetime}, {"next_service_date", type datetime}, {"certificate_number", type text}, {"wire_roll_cert_number", type text}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"type"}, WireRanges, {"Type"}, "WireRanges", JoinKind.LeftOuter),
    #"Expanded WireRanges" = Table.ExpandTableColumn(#"Merged Queries", "WireRanges", {"Min", "Max"}, {"Min", "Max"})
in
    #"Expanded WireRanges"