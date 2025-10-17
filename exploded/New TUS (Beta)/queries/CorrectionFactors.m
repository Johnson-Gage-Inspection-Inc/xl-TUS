let
    // 1. Load and expand WireSets into individual SNs (with defensive loading)
    WireSetsRaw = try Excel.CurrentWorkbook(){[Name="WireSets"]}[Content] otherwise #table({}, {}),

    // Early exit check: if critical tables are empty, return empty result
    IsWorkbookBlank = Table.RowCount(WireSetsRaw) = 0,

    ExpandedWireSets = if IsWorkbookBlank then #table({}, {}) else Table.AddColumn(WireSetsRaw, "wire_sn", each
    let
        serial = [serial_number],
        parts = Text.Split(serial, "-")
    in
        if List.Count(parts) = 2 then
            let
                prefix = Text.Select(parts{0}, {"A".."Z"}),
                startNum = Number.FromText(Text.Select(parts{0}, {"0".."9"})),
                endNum = Number.FromText(Text.Select(parts{1}, {"0".."9"}))
            in
                List.Transform({startNum..endNum}, each prefix & Text.PadStart(Text.From(_), 2, "0"))
        else
            {}  // fallback: empty list for unparsable rows
, type list),
    FlattenedWireSets = if IsWorkbookBlank then #table({}, {}) else Table.ExpandListColumn(ExpandedWireSets, "wire_sn"),
    RollMap = if IsWorkbookBlank then #table({}, {}) else try Table.SelectColumns(FlattenedWireSets, {"wire_roll_cert_number", "wire_sn"}) otherwise #table({"wire_roll_cert_number", "wire_sn"}, {}),

    // 2. Load used SNs from Main!O5:O44 (with defensive loading)
    UsedWireSNs = try Excel.CurrentWorkbook(){[Name="UsedWireSNs"]}[Content] otherwise #table({}, {}),
    #"Renamed Columns1" = Table.RenameColumns(UsedWireSNs,{{"Column1", "Thermocouple"}}),    WithIndex = if IsWorkbookBlank then #table({}, {}) else Table.AddIndexColumn(#"Renamed Columns1", "Index", 1, 1, Int64.Type),
    NonNulls = if IsWorkbookBlank then #table({}, {}) else Table.SelectRows(WithIndex, each [Thermocouple] <> null),

    #"Reordered Columns" = Table.ReorderColumns(NonNulls,{"Index", "Thermocouple"}),

    // 5. Look up wire lot numbers
    Joined = if IsWorkbookBlank then #table({}, {}) else Table.NestedJoin(#"Reordered Columns", {"Thermocouple"}, RollMap, {"wire_sn"}, "Match", JoinKind.LeftOuter),
    Expanded = if IsWorkbookBlank then #table({}, {}) else
        if List.Contains(Table.ColumnNames(Joined), "Match") then
            try Table.ExpandTableColumn(Joined, "Match", {"wire_roll_cert_number"}) otherwise Table.AddColumn(Joined, "wire_roll_cert_number", each null, type text)
        else
            #table({}, {}),

    // 6. Interpolate wire offsets
    TestPoints = Table.RenameColumns(Expanded,{{"wire_roll_cert_number", "WireLotNumber"}}),
    // 1. Get the test temperature
    TestTemp = try Number.From(Excel.CurrentWorkbook(){[Name="TestTemp"]}[Content]{0}[Column1]) otherwise null,
    #"Reordered Columns2" = Table.ReorderColumns(TestPoints,{"Index", "Thermocouple"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns2",{{"Index", "point"}}),

    // 2. Interpolate daqbook offsets
    DaqbookTable = Excel.CurrentWorkbook(){[Name="DaqbookServiceRecord"]}[Content],
    RawTN = try DaqbookTable{11}[Value] otherwise "", // Get the TN like "J31124" (with defensive loading)
    TN_WithUnderscore = if Text.Length(RawTN) >= 4 then Text.Start(RawTN, 2) & "_" & Text.Middle(RawTN, 2) else "", // Convert to filename "J3_1124"
    AllDaqbookOffsetsCached = Excel.CurrentWorkbook(){[Name="AllDaqbookOffsets"]}[Content],
    #"Changed Type1" = Table.TransformColumnTypes(AllDaqbookOffsetsCached,{{"Offset", type number}, {"Temp", type number}, {"reading", type number}, {"traceability_no", type text}, {"point", Int64.Type}}),
    // Force cache of the upstream table (with defensive loading)
    WithDaqbookOffsets = Table.AddColumn(#"Renamed Columns", "DaqbookOffset", each LinearInterpolate(#"Changed Type1", TestTemp, TN_WithUnderscore, [point])),
    AllWireOffsetsCached = Excel.CurrentWorkbook(){[Name="AllWireOffsets"]}[Content],
    #"Changed Type2" = Table.TransformColumnTypes(AllWireOffsetsCached,{{"Temp", type number}, {"Offset", type number}}),
    WithTCOffsets = Table.AddColumn(WithDaqbookOffsets, "TCOffset", each LinearInterpolate(#"Changed Type2", TestTemp, [WireLotNumber])),
    #"Added Custom" = Table.AddColumn(WithTCOffsets, "CummulativeOffset", each [DaqbookOffset]+[TCOffset]),
    #"Changed Type" = Table.TransformColumnTypes(#"Added Custom",{{"DaqbookOffset", type number}, {"TCOffset", type number}, {"CummulativeOffset", type number}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Changed Type",{"point", "Thermocouple", "WireLotNumber", "DaqbookOffset", "TCOffset", "CummulativeOffset"})
in
    #"Reordered Columns1"