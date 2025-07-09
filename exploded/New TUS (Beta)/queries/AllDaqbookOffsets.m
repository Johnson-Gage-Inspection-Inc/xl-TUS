let
    Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "daqbook-offsets/"])),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"point", "reading", "temp", "tn"}, {"point", "reading", "temp", "tn"}),
    #"Added Custom" = Table.AddColumn(#"Expanded Column1", "Custom", each [reading]-[temp]),
    #"Renamed Columns" = Table.RenameColumns(#"Added Custom",{{"Custom", "Offset"}, {"temp", "Temp"}, {"tn", "traceability_no"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Offset", type number}, {"Temp", type number}, {"reading", type number}, {"traceability_no", type text}})
in
    #"Changed Type"