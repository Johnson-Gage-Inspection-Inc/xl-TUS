let
    // Try to get fresh data from SharePoint
    FreshData = try 
        let
            Source = Excel.Workbook(Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared Documents/Pyro/ClientOvens.xlsx"), null, true),
            Header_Info_Table = Source{[Item="Header_Info",Kind="Table"]}[Data],
            #"Changed Type1" = Table.TransformColumnTypes(Header_Info_Table,{{"AssetId", Int64.Type}, {"ClientCompanyId", Int64.Type}, {"Furnace", type any}, {"FormNumber", type text}, {"CalibrationMethod", type text}, {"ToleranceSource", type text}, {"Item", type text}, {"ModelNumber", type any}, {"SerialNumber", type any}, {"UnitNumber", type any}, {"Class", type any}, {"HeatingMethod", type text}, {"WorkingZoneSize", type text}, {"CubicFeet", type any}, {"CalInterval", type any}, {"ControllerId", type any}, {"Controller", type any}, {"ContSN", type any}, {"ContTol", type any}, {"RecorderId", type any}, {"Recorder", type any}, {"RecSN", type any}, {"RecTol", type any}, {"FurnaceSpecificComments", type text}, {"TestLocation", type text}, {"OvenLocation", type text}, {"Condition", type text}, {"PB", type any}, {"Integral", type any}, {"Derivative", type any}, {"Other1", type any}, {"Other2", type any}, {"CheckRec", type any}, {"Load", type text}, {"LagLimit", type number}, {"RecoveryLimit", type number}}),
            #"Changed Type" = Table.TransformColumnTypes(#"Changed Type1",{{"AssetId", Int64.Type}, {"ClientCompanyId", Int64.Type}, {"Furnace", type any}, {"FormNumber", type text}, {"CalibrationMethod", type text}, {"ToleranceSource", type text}, {"Item", type text}, {"ModelNumber", type any}, {"SerialNumber", type any}, {"UnitNumber", type any}, {"Class", type any}, {"HeatingMethod", type text}, {"WorkingZoneSize", type text}, {"CubicFeet", type any}, {"CalInterval", type any}, {"ControllerId", type any}, {"Controller", type any}, {"ContSN", type any}, {"ContTol", type any}, {"RecorderId", type any}, {"Recorder", type any}, {"RecSN", type any}, {"RecTol", type any}, {"FurnaceSpecificComments", type text}, {"TestLocation", type text}, {"OvenLocation", type text}, {"Condition", type text}, {"PB", type any}, {"Integral", type any}, {"Derivative", type any}, {"Other1", type any}, {"Other2", type any}, {"CheckRec", type any}, {"Load", type text}, {"LagLimit", type number}, {"RecoveryLimit", type number}}),
            #"Clean" = Table.ReplaceErrorValues(#"Changed Type", {{"ModelNumber", null}, {"SerialNumber", null}, {"Class", null}, {"CubicFeet", null}, {"CalInterval", null}, {"ControllerId", null}, {"Controller", null}, {"ContSN", null}, {"ContTol", null}, {"RecorderId", null}, {"Recorder", null}, {"RecSN", null}, {"RecTol", null}, {"TestLocation", null}, {"PB", null}, {"Derivative", null}, {"Integral", null}, {"Other1", null}, {"Other2", null}, {"CheckRec", null}}),
            #"Filtered Rows" = Table.SelectRows(#"Clean", each ([AssetId] <> null))
        in
            #"Filtered Rows"
    otherwise null,
    
    // Try to get cached data from workbook
    CachedData = try Excel.CurrentWorkbook(){[Name="Header_Info"]}[Content] otherwise null,
    
    // Use fresh data if available, otherwise use cached data, otherwise return empty table
    FinalData = if FreshData <> null then FreshData 
                else if CachedData <> null then CachedData
                else #table(
                    {"AssetId", "ClientCompanyId", "Furnace", "FormNumber", "CalibrationMethod", "ToleranceSource", "Item", "ModelNumber", "SerialNumber", "UnitNumber", "Class", "HeatingMethod", "WorkingZoneSize", "CubicFeet", "CalInterval", "ControllerId", "Controller", "ContSN", "ContTol", "RecorderId", "Recorder", "RecSN", "RecTol", "FurnaceSpecificComments", "TestLocation", "OvenLocation", "Condition", "PB", "Integral", "Derivative", "Other1", "Other2", "CheckRec", "Load", "LagLimit", "RecoveryLimit"},
                    {{0, 0, "No data available", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 0, 0}}
                )
in
    FinalData