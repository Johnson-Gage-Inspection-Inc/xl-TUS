let
    Source = CorrectionFactors,
    #"Removed Other Columns" = Table.SelectColumns(Source,{"WireLotNumber", "TCOffset"}),
    #"Removed Duplicates" = Table.Distinct(#"Removed Other Columns")
in
    #"Removed Duplicates"