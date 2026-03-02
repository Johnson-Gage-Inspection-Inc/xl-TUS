let
    Source = Json.Document(Web.Contents("https://jgiapi.com/daqbook-service-records/latest/")),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"asset_id", "certificate_tn", "due_date", "service_date"}, {"asset_id", "certificate_tn", "due_date", "service_date"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Column1",{{"service_date", type datetime}, {"due_date", type datetime}, {"asset_id", Int64.Type}, {"certificate_tn", type text}}),
    #"Inserted Text Before Delimiter" = Table.AddColumn(#"Changed Type", "asset_tag", each Text.BeforeDelimiter([certificate_tn], "_"), type text),
    #"Reordered Columns" = Table.ReorderColumns(#"Inserted Text Before Delimiter",{"asset_tag", "asset_id", "certificate_tn", "due_date", "service_date"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns",{{"service_date", type date}, {"due_date", type date}})
in
    #"Changed Type1"