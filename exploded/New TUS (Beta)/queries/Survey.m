let
    // Get selection from radio button: 1 = PTs 1–14, 2 = PTs 15–28, 3 = PTs 29–40
    RadioSelection = try Excel.CurrentWorkbook(){[Name="PT_Range_Selection"]}[Content]{0}[Column1] otherwise 1,

    // Resolve channel range based on selection
    ChannelRange =
        if RadioSelection = 2 then [First=15, Last=28]
        else if RadioSelection = 3 then [First=29, Last=40]
        else [First=1, Last=14],

    FIRST_CHANNEL = ChannelRange[First],
    LAST_CHANNEL = ChannelRange[Last],

    SurveyStartTime = Time.From(Excel.CurrentWorkbook(){[Name="TUS_Start_Time"]}[Content]{0}[Column1]),
    SurveyEndTime = Time.From(Excel.CurrentWorkbook(){[Name="Survey_End_Time"]}[Content]{0}[Column1]),
    Source = Excel.CurrentWorkbook(){[Name="DataForChannels1to14"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Time", type time}, {"1", type number}, {"2", type number}, {"3", type number}, {"4", type number}, {"5", type number}, {"6", type number}, {"7", type number}, {"8", type number}, {"9", type number}, {"10", type number}, {"11", type number}, {"12", type number}, {"13", type number}, {"14", type number}}),
    Source2 = Excel.CurrentWorkbook(){[Name="DataForChannels15to28"]}[Content],
    DataForChannels15to28 = Table.TransformColumnTypes(Source2,{{"Time", type time}, {"15", type number}, {"16", type number}, {"17", type number}, {"18", type number}, {"19", type number}, {"20", type number}, {"21", type number}, {"22", type number}, {"23", type number}, {"24", type number}, {"25", type number}, {"26", type number}, {"27", type number}, {"28", type number}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Time"}, DataForChannels15to28, {"Time"}, "DataForChannels15to28", JoinKind.FullOuter),
    #"Expanded DataForChannels15to28" = Table.ExpandTableColumn(#"Merged Queries", "DataForChannels15to28", {"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28"}, {"15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28"}),
    Source3 = Excel.CurrentWorkbook(){[Name="DataForChannels29to40"]}[Content],
    DataForChannels29to40 = Table.TransformColumnTypes(Source3,{{"Time", type time}, {"29", type number}, {"30", type number}, {"31", type number}, {"32", type number}, {"33", type number}, {"34", type number}, {"35", type number}, {"36", type number}, {"37", type number}, {"38", type number}, {"39", type number}, {"40", type number}}),
    #"Merged Queries1" = Table.NestedJoin(#"Expanded DataForChannels15to28", {"Time"}, DataForChannels29to40, {"Time"}, "DataForChannels29to40", JoinKind.FullOuter),
    #"Expanded DataForChannels29to40" = Table.ExpandTableColumn(#"Merged Queries1", "DataForChannels29to40", {"29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40"}, {"29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40"}),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Expanded DataForChannels29to40", {"Time"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Attribute", "TestPoint"}, {"Value", "RawTemp"}}),
    #"Filtered Rows" = Table.SelectRows(#"Renamed Columns", each [Time] >= SurveyStartTime and [Time] <= SurveyEndTime),
    #"Changed Type1" = Table.TransformColumnTypes(#"Filtered Rows",{{"TestPoint", Int64.Type}}),
    // Load CorrectionFactors table with comprehensive error handling
    CorrectionFactors = try Excel.CurrentWorkbook(){[Name="CorrectionFactors"]}[Content] otherwise #table({"point", "CummulativeOffset", "TCOffset", "DaqbookOffset"}, {}),
    #"Changed Type2" = try (
        if Table.IsEmpty(CorrectionFactors) then
            Table.TransformColumnTypes(CorrectionFactors,{{"point", Int64.Type}, {"CummulativeOffset", type number}, {"TCOffset", type number}, {"DaqbookOffset", type number}})
        else
            Table.TransformColumnTypes(CorrectionFactors,{{"CummulativeOffset", type number}, {"TCOffset", type number}, {"DaqbookOffset", type number}})
    ) otherwise Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(#table({}, {}), "point", each null, Int64.Type), "CummulativeOffset", each null, type number), "TCOffset", each null, type number), "DaqbookOffset", each null, type number),
    // Join on TestPoint to get CummulativeOffset
    #"Merged with CF" = Table.NestedJoin(#"Changed Type1", {"TestPoint"}, #"Changed Type2", {"point"}, "CF", JoinKind.LeftOuter),
    #"Expanded CF" = try Table.ExpandTableColumn(#"Merged with CF", "CF", {"CummulativeOffset"}) otherwise Table.AddColumn(#"Merged with CF", "CummulativeOffset", each null, type number),

    // Add CorrectedTemp = RawTemp + CummulativeOffset
    #"Added CorrectedTemp" = Table.AddColumn(#"Expanded CF", "CorrectedTemp", each [RawTemp] + [CummulativeOffset], type number),

    // Define empty fallback with correct column types and names
    EmptySchema = #table(
        {"Time", "TestPoint", "RawTemp", "CummulativeOffset", "CorrectedTemp"},
        {}
    ),

    // Force final structure even if no data remains
    Final = Table.Combine({EmptySchema, #"Added CorrectedTemp"}),
    #"Changed Type3" = Table.TransformColumnTypes(Final,{{"Time", type time}, {"TestPoint", Int64.Type}, {"RawTemp", type number}, {"CummulativeOffset", type number}, {"CorrectedTemp", type number}}),
    #"Filtered Rows1" = Table.SelectRows(#"Changed Type3", each [TestPoint] >= FIRST_CHANNEL and [TestPoint] <= LAST_CHANNEL)
in
    #"Filtered Rows1"