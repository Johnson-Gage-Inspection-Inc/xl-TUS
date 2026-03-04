// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    True

let
    Source = AllWireOffsets,
    Grouped = Table.Group(
        Source,
        {"traceability_no", "wire_type"},
        {{"MaxTemp", each List.Max([Temp]), type nullable number}}
    ),

    WithClass = Table.AddColumn(
        Grouped,
        "TempClass",
        each
            let t = [MaxTemp]
            in
                if t <> null and t > 2000 then 3
                else if t <> null and t > 500 then 2
                else 1,
        Int64.Type
    )
in
    WithClass