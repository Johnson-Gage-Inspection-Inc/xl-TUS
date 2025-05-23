let
    Status = try Text.From(Excel.CurrentWorkbook(){[Name="Status"]}[Content][Column1]{0}) otherwise "",
    CompanyId = try Text.From(Excel.CurrentWorkbook(){[Name="CompanyId"]}[Content][Column1]{0}) otherwise "",
    From = try Text.From(Excel.CurrentWorkbook(){[Name="From"]}[Content][Column1]{0}) otherwise "",
    To = try Text.From(Excel.CurrentWorkbook(){[Name="To"]}[Content][Column1]{0}) otherwise "",
    WorkItemNumber = try Text.From(Excel.CurrentWorkbook(){[Name="WorkItemNumber"]}[Content][Column1]{0}) otherwise "",
    AssetSearch = try Text.From(Excel.CurrentWorkbook(){[Name="AssetSearch"]}[Content][Column1]{0}) otherwise "",
    status = if Text.Length(Status) > 0 then [ status = Status ] else [],
    companyId = if Text.Length(CompanyId) > 0 then [ companyId = CompanyId ] else [],
    from = if Text.Length(From) > 0 then [ from = From ] else [],
    to = if Text.Length(To) > 0 then [ to = To ] else [],
    workItemNumber = if Text.Length(WorkItemNumber) > 0 then [ workItemNumber = WorkItemNumber ] else [],
    assetSearch = if Text.Length(AssetSearch) > 0 then [ assetSearch = AssetSearch ] else [],
    QueryOptions = Record.Combine({status, companyId, from, to, workItemNumber, assetSearch}),
    baseUrl = "https://jgiquality.qualer.com",
    relativeUrl = "api/service/workitems",
    response = Web.Contents(
        baseUrl,
        [
            RelativePath = relativeUrl,
            Query = QueryOptions,
            Headers = [ Authorization = TokenText ]
        ]
    ),
    json = Json.Document(response),
    ConvertToTable = Table.FromRecords(json)
in
    ConvertToTable