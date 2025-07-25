(table as table, targetTemp as number, optional matchValue as nullable text, optional matchIndex as nullable number) as nullable number =>
let
    // First: optional filter by matchValue (WireLotNumber or TraceabilityNumber)
    filtered1 = 
        if matchValue <> null and Table.HasColumns(table, "WireLotNumber") then
            Table.SelectRows(table, each [WireLotNumber] = matchValue)
        else if matchValue <> null and Table.HasColumns(table, "TraceabilityNumber") then
            Table.SelectRows(table, each [TraceabilityNumber] = matchValue)
        else
            table,

    // Second: optional filter by Index (channel number)
    filtered2 = 
        if matchIndex <> null and Table.HasColumns(filtered1, "point") then
            Table.SelectRows(filtered1, each [point] = matchIndex)
        else
            filtered1,

    // Defensive fallback if table is empty or columns are missing
    safeTable = 
        if Table.HasColumns(filtered2, "Temp") and Table.HasColumns(filtered2, "Offset") and Table.RowCount(filtered2) > 0 then
            filtered2
        else
            #table({"Temp", "Offset"}, {}),

    sorted = Table.Sort(safeTable, {{"Temp", Order.Ascending}}),
    below = Table.SelectRows(sorted, each [Temp] <= targetTemp),
    above = Table.SelectRows(sorted, each [Temp] >= targetTemp),
    lower = if Table.RowCount(below) > 0 then List.Last(below[Temp]) else null,
    upper = if Table.RowCount(above) > 0 then List.First(above[Temp]) else null,
    lowerOffset = if lower <> null then Record.Field(Table.SelectRows(sorted, each [Temp] = lower){0}, "Offset") else null,
    upperOffset = if upper <> null then Record.Field(Table.SelectRows(sorted, each [Temp] = upper){0}, "Offset") else null,
    rawResult = 
        if lower = null or upper = null or lower = upper then lowerOffset
        else lowerOffset + (targetTemp - lower) * (upperOffset - lowerOffset) / (upper - lower),
    result = if rawResult <> null then Number.Round(rawResult, 1, RoundingMode.ToEven) else null
in
    result