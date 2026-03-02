let
    Source = AllWireOffsets,
    Grouped = Table.Group(Source, {"traceability_no", "wire_type"}, {{"MaxTemp", each List.Max([Temp]), type nullable number}}),

    WithClass =
        Table.AddColumn(
            Grouped,
            "TempClass",
            each
                let t = [MaxTemp]
                in
                    if t <> null and t > 2000 then "Ultra High Temp"
                    else if t <> null and t > 500 then "High Temp"
                    else null,
            type text
        )
in
    WithClass