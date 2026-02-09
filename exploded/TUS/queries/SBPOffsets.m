let
    response = Json.Document(Web.Contents("https://jgiapi.com/sbp-offsets/")),
    #"Converted to Table" = Table.FromList(response, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"certificate_number", "nominal_temp", "offset", "probe_serial_number"}, {"certificate_number", "nominal_temp", "offset", "probe_serial_number"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Column1",{{"probe_serial_number", type text}, {"certificate_number", type text}, {"nominal_temp", type number}, {"offset", type number}})
in
    #"Changed Type"