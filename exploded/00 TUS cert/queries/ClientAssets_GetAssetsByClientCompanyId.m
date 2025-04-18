let
    ClientCompanyId = Text.From(Excel.CurrentWorkbook(){[Name="ClientCompanyId"]}[Content][Column1]{0}),
    Query_equipmentId = try Text.From(Excel.CurrentWorkbook(){[Name="Query.equipmentId"]}[Content][Column1]{0}) otherwise "",
    Query_serialNumber = try Text.From(Excel.CurrentWorkbook(){[Name="Query.serialNumber"]}[Content][Column1]{0}) otherwise "",
    Query_assetTag = try Text.From(Excel.CurrentWorkbook(){[Name="Query.assetTag"]}[Content][Column1]{0}) otherwise "",
    Query_barcode = try Text.From(Excel.CurrentWorkbook(){[Name="Query.barcode"]}[Content][Column1]{0}) otherwise "",
    Query_legacyId = try Text.From(Excel.CurrentWorkbook(){[Name="Query.legacyId"]}[Content][Column1]{0}) otherwise "",
    query.equipmentId = if Text.Length(Query_equipmentId) > 0 then [ query.equipmentId = Query_equipmentId ] else [],
    query.serialNumber = if Text.Length(Query_serialNumber) > 0 then [ query.serialNumber = Query_serialNumber ] else [],
    query.assetTag = if Text.Length(Query_assetTag) > 0 then [ query.assetTag = Query_assetTag ] else [],
    query.barcode = if Text.Length(Query_barcode) > 0 then [ query.barcode = Query_barcode ] else [],
    query.legacyId = if Text.Length(Query_legacyId) > 0 then [ query.legacyId = Query_legacyId ] else [],
    QueryOptions = Record.Combine({query.equipmentId, query.serialNumber, query.assetTag, query.barcode, query.legacyId}),
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/clients/" & ClientCompanyId & "/assets",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = Text.TrimStart(relativeUrl, "/"),
            Query = QueryOptions,
            Headers = [ Authorization = "Api-Token REDACTED" ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json)
in
    ConvertToTable