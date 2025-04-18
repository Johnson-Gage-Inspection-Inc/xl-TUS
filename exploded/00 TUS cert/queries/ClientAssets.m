let
    ConvertToTable = ClientAssets_GetAssetsByClientCompanyId,
    #"Filtered Rows" = Table.SelectRows(ConvertToTable, each [RootCategoryName] = "Furnaces" or [CategoryName] = "Furnaces")
in
    #"Filtered Rows"