// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    False

let
    RequiredColumns = {
        "id",
        "sync_type",
        "category",
        "severity",
        "asset_tag",
        "asset_id",
        "short_message",
        "detail_hint",
        "first_seen_at",
        "last_seen_at"
    },

    EmptyAlerts = #table(RequiredColumns, {}),

    FreshData =
        try
            let
                Response = Web.Contents(
                    "https://jgiapi.com",
                    [
                        RelativePath = "data-sync/alerts",
                        Query = [limit = "250"],
                        ManualStatusHandling = {200, 204, 400, 401, 403, 404, 409, 422, 500},
                        Timeout = #duration(0, 0, 0, 30)
                    ]
                ),
                Status = try Value.Metadata(Response)[Response.Status] otherwise 200,
                ParsedJson = try Json.Document(Response) otherwise null,
                Json = if Status = 200 then ParsedJson else null,
                AsTable =
                    if Json <> null and Value.Is(Json, type list) then
                        try Table.FromRecords(Json) otherwise null
                    else
                        null,
                Normalized =
                    if AsTable = null then
                        null
                    else
                        Table.SelectColumns(AsTable, RequiredColumns, MissingField.UseNull)
            in
                Normalized
        otherwise
            null,

    CachedData =
        try
            let
                cached = Excel.CurrentWorkbook(){[Name="SyncAlerts"]}[Content],
                normalized = Table.SelectColumns(cached, RequiredColumns, MissingField.UseNull)
            in
                if Table.RowCount(normalized) > 0 then normalized else null
        otherwise
            null,

    FinalData =
        if FreshData <> null then FreshData
        else if CachedData <> null then CachedData
        else EmptyAlerts
in
    FinalData