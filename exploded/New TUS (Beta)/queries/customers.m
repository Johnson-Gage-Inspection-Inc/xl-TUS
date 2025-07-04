let
    response = Web.Contents("https://api.jgiquality.com", [RelativePath = "clients"]),
    json = Json.Document(response),
    ConvertToTable = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ExpandRecords = Table.ExpandRecordColumn(ConvertToTable, "Column1", Record.FieldNames(ConvertToTable{0}[Column1])),
    #"Removed Other Columns" = Table.SelectColumns(ExpandRecords,{"company_id", "company_name", "shipping_address"}),
    #"Expanded shipping_address" = Table.ExpandRecordColumn(#"Removed Other Columns", "shipping_address", {"additional_properties"}, {"additional_properties"}),
    #"Expanded additional_properties" = Table.ExpandRecordColumn(#"Expanded shipping_address", "additional_properties", {"Address1", "Address2", "Attention1", "Attention2", "City", "Company", "Country", "Email", "FaxNumber", "FirstName", "LastName", "PhoneNumber", "StateProvince", "StateProvinceAbbreviation", "ZipPostalCode"}, {"Address1", "Address2", "Attention1", "Attention2", "City", "Company", "Country", "Email", "FaxNumber", "FirstName", "LastName", "PhoneNumber", "StateProvince", "StateProvinceAbbreviation", "ZipPostalCode"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded additional_properties",{"company_name", "Address1", "Address2", "City", "StateProvinceAbbreviation", "ZipPostalCode", "company_id"}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Reordered Columns",{{"City", Text.Proper, type text}, {"Address1", Text.Proper, type text}, {"Address2", Text.Proper, type text}}),
    #"Replaced Value1" = Table.ReplaceValue(#"Capitalized Each Word",null,"KS",Replacer.ReplaceValue,{"StateProvinceAbbreviation"}),
    #"Merged Columns" = Table.CombineColumns(#"Replaced Value1",{"City", "StateProvinceAbbreviation", "ZipPostalCode"},Combiner.CombineTextByDelimiter(", ", QuoteStyle.None),"Merged"),
    #"Replaced Value" = Table.ReplaceValue(#"Merged Columns",null,"",Replacer.ReplaceValue,{"Address2"}),
    #"Sorted Rows" = Table.Sort(#"Replaced Value",{{"company_name", Order.Ascending}})
in
    #"Sorted Rows"