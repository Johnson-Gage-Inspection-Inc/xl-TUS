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
                Source = Json.Document(Web.Contents(
                    "https://jgiapi.com",
                    [
                        RelativePath = "data-sync/alerts",
                        Query = [limit = "250"]
                    ]
                )),
                AsTable =
                    if Value.Is(Source, type list) then
                        try Table.FromRecords(Source) otherwise null
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
                Existing = Excel.CurrentWorkbook(){[Name="Alerts"]}[Content],
                NormalizedCached = Table.SelectColumns(Existing, RequiredColumns, MissingField.UseNull)
            in
                NormalizedCached
        otherwise
            null,

    FinalData =
        if FreshData <> null then FreshData
        else if CachedData <> null then CachedData
        else EmptyAlerts
in
    FinalData