let
    Source = Excel.CurrentWorkbook(){[Name="Comparison_Report_Data"]}[Content],
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Renamed Columns2" = Table.RenameColumns(#"Promoted Headers",{{"UUT Resolution", "Resolution"}}),
    #"Added Index" = Table.AddIndexColumn(#"Renamed Columns2", "Index", 1, 1, Int64.Type),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Added Index", {"Index", "Resolution", "Item Name", "TUS Point #"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "Time"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"TUS Point #", Int64.Type}, {"Item Name", type text}}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Changed Type",{{"Time", type number}, {"Value", type number}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Changed Type3",{{"Time", type time}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type2", {"Time", "TUS Point #"}, Survey, {"Time", "TestPoint"}, "Survey", JoinKind.LeftOuter),
    #"Expanded Survey" = Table.ExpandTableColumn(#"Merged Queries", "Survey", {"CorrectedTemp"}, {"CorrectedTemp"}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Expanded Survey",{{"Resolution", Int64.Type}}),
    #"Rounded Value" = Table.AddColumn(#"Changed Type4", "RoundedValue", each Number.Round([Value], [Resolution]), type number),
    #"Rounded Temp" = Table.AddColumn(#"Rounded Value", "RoundedCorrectedTemp", each Number.Round([CorrectedTemp], [Resolution]), type number),
    #"Removed Unrounded" = Table.RemoveColumns(#"Rounded Temp", {"Value", "CorrectedTemp"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Removed Unrounded",{{"RoundedValue", "Value"}, {"RoundedCorrectedTemp", "CorrectedTemp"}}),
    #"Added Custom" = Table.AddColumn(#"Renamed Columns1", "Deviation", each [Value] - [CorrectedTemp])
in
    #"Added Custom"