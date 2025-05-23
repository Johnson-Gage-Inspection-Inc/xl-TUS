let
    // Get the TN like "J31124"
    RawTN = Excel.CurrentWorkbook(){[Name="TN"]}[Content]{0}[Column1],
    // Convert to filename "J3_1124"
    TN_WithUnderscore = Text.Start(RawTN, 2) & "_" & Text.Middle(RawTN, 2),
    // Build SharePoint URL
    FileUrl = "https://jgiquality.sharepoint.com/sites/JGI/Shared%20Documents/Pyro/Pyro_Standards/" & TN_WithUnderscore & ".xlsm",

    // Load workbook and get Sheet1
    Workbook = Excel.Workbook(Web.Contents(FileUrl), null, true),
    Sheet1 = Workbook{[Item="Sheet1",Kind="Sheet"]}[Data],

    // Extract test temps from A42:A47
    Temps = List.Transform({42..47}, each Record.Field(Sheet1{_ - 1}, "Column1")),

    // Generate row offsets for point blocks (like in VBA)
    BlockOffsets = {42, 50, 60, 68, 78, 86, 96},
    ChannelCols = {"Column2", "Column3", "Column4", "Column5", "Column6", "Column7"},

    // Flatten: For each temp (i), for each point row (r), for each channel (c)
    Output = List.Combine(
        List.Transform({0..List.Count(Temps)-1}, (i) =>
            let
                TempValue = Temps{i},
                RowsForThisTemp = List.Combine(
                    List.Transform({0..6}, (b) =>
                        let
                            RowOffset = BlockOffsets{b} + i,
                            BaseRow = Sheet1{RowOffset - 1},
                            Values = List.Transform(ChannelCols, (col) => Record.Field(BaseRow, col)),
                            Records = List.Transform({0..List.Count(Values)-1}, (c) =>
                                [Temp=TempValue, Point = (b * 6) + c + 1, RawValue=Values{c}, Delta=Number.Round((Values{c} - TempValue) * -1, 2)]
                            )
                        in
                            Records
                    )
                )
            in
                RowsForThisTemp
        )
    ),

    // Convert to table
    OutputTable = Table.FromRecords(Output),
    #"Filtered Rows" = Table.SelectRows(OutputTable, each ([RawValue] <> null)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"RawValue"}),
    #"Sorted Rows" = Table.Sort(#"Removed Columns",{{"Point", Order.Ascending}, {"Temp", Order.Ascending}}),
    // Group by Point
    Grouped = Table.Group(#"Sorted Rows", {"Point"}, {
        {"Expanded", each
            let
                rows = _,
                sorted = Table.Sort(rows, {{"Temp", Order.Ascending}}),
                withMidpoints = List.Combine(
                    List.Transform({0..Table.RowCount(sorted)-2}, (i) =>
                        let
                            row1 = sorted{i},
                            row2 = sorted{i+1},
                            midTemp = Number.Round((row1[Temp] + row2[Temp]) / 2, 0),
                            midDelta = Number.Round((row1[Delta] + row2[Delta]) / 2, 2),
                            midpointRecord = [Point=row1[Point], Temp=midTemp, Delta=midDelta]
                        in
                            {row1, midpointRecord}
                    ) & { {sorted{Table.RowCount(sorted)-1}} }  // include last row
                )
            in
                withMidpoints,
            type list
        }
    }),

    // Expand again
    #"Expanded Rows" = Table.ExpandListColumn(Grouped, "Expanded"),
    #"Expanded Records" = Table.ExpandRecordColumn(#"Expanded Rows", "Expanded", {"Temp", "Delta"}),
    // Questinon: Is this bankers rounding? i.e. round to even?
    #"Rounded Off" = Table.TransformColumns(#"Expanded Records",{{"Delta", each Number.Round(_, 1), type number}}),

    // Pivot on new expanded data with midpoints
    Pivoted = Table.Pivot(
        Table.TransformColumnTypes(#"Rounded Off", {{"Temp", type text}}, "en-US"),
        List.Distinct(Table.TransformColumnTypes(#"Rounded Off", {{"Temp", type text}}, "en-US")[Temp]),
        "Temp",
        "Delta",
        List.Sum
    )
in
    Pivoted