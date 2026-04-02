// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    False

let
  // Force Sync to run first
  _ = try Table.RowCount(Sync) otherwise null,

  // Then do the GET
  Source = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath="wire-set-certs/"])),
  L = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
  T = Table.ExpandRecordColumn(L, "Column1",
        {"asset_id","asset_tag","certificate_number","serial_number","service_date",
         "traceability_number","wire_roll_cert_number","wire_set_group"},
        {"asset_id","asset_tag","certificate_number","serial_number","service_date",
         "traceability_number","wire_roll_cert_number","wire_set_group"}),
  Out = Table.Sort(T, {{"asset_tag", Order.Ascending}})
in
  Out