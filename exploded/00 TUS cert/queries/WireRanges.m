let
    Source = Excel.CurrentWorkbook(){[Name="WireRanges"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Type", type text}, {"Min", Int64.Type}, {"Max", Int64.Type}})
in
    #"Changed Type"