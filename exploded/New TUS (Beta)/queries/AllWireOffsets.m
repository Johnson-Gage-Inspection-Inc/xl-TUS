let
    Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath = "wire-offsets/"])),
    #"Converted to Table" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"correction_factor", "created_at", "id", "nominal_temp", "traceability_no", "updated_at", "updated_by"}, {"correction_factor", "created_at", "id", "nominal_temp", "traceability_no", "updated_at", "updated_by"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Expanded Column1",{"traceability_no", "nominal_temp", "correction_factor"}),
    #"Renamed Columns" = Table.RenameColumns(#"Removed Other Columns",{{"correction_factor", "Offset"}, {"nominal_temp", "Temp"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"Offset", type number}, {"Temp", type number}, {"traceability_no", type text}})
in
    #"Changed Type"