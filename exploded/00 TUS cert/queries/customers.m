let
    Source = Json.Document(
        Web.Contents(
            "https://jgiquality.qualer.com/api/service/clients",
            [
                Headers = [
                    #"Authorization" = "Api-Token bf407589-f463-4046-ba2c-30642bd5d637"
                ]
            ]
        )
    ),
    ConvertToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ExpandRecords = Table.ExpandRecordColumn(ConvertToTable, "Column1", Record.FieldNames(ConvertToTable{0}[Column1])),
    #"Removed Other Columns" = Table.SelectColumns(ExpandRecords,{"CompanyId", "CompanyName", "ShippingAddress"}),
    #"Expanded ShippingAddress" = Table.ExpandRecordColumn(#"Removed Other Columns", "ShippingAddress", {"City", "Address1", "Address2", "ZipPostalCode", "StateProvinceAbbreviation"}, {"City", "Address1", "Address2", "ZipPostalCode", "StateProvinceAbbreviation"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded ShippingAddress",{"CompanyName", "Address1", "Address2", "City", "StateProvinceAbbreviation", "ZipPostalCode", "CompanyId"}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Reordered Columns",{{"City", Text.Proper, type text}, {"Address1", Text.Proper, type text}, {"Address2", Text.Proper, type text}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Capitalized Each Word",null,"KS",Replacer.ReplaceValue,{"StateProvinceAbbreviation"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced Value1",{"City", "StateProvinceAbbreviation", "ZipPostalCode"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Merged"),
    #"Replaced Value" = Table.ReplaceValue(#"Merged Columns",null,"",Replacer.ReplaceValue,{"Address2"}),
    #"Sorted Rows" = Table.Sort(#"Replaced Value",{{"CompanyName", Order.Ascending}})
in
    #"Sorted Rows"