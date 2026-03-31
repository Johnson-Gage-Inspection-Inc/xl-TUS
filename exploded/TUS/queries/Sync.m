// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    False

let
  // 1) Table users edit (must be named WireSetCerts)
  SheetTable =
    try Excel.CurrentWorkbook(){[Name="WireSetCerts"]}[Content]
    otherwise error "No table named 'WireSetCerts' found.",

  // Keep only what we need
  Sheet = Table.SelectColumns(SheetTable, {"asset_id","wire_roll_cert_number"}, MissingField.UseNull),

  // 2) Server truth (only the same two fields)
  ApiJson = Json.Document(Web.Contents("https://jgiapi.com", [RelativePath="wire-set-certs/"])),
  ApiListTable = Table.FromList(ApiJson, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
  ApiExpanded = Table.ExpandRecordColumn(
                  ApiListTable, "Column1",
                  {"asset_id","wire_roll_cert_number"},
                  {"asset_id","server_wire_roll_cert_number"}),

  // 3) Join by asset_id only
  Joined = Table.NestedJoin(Sheet, {"asset_id"}, ApiExpanded, {"asset_id"}, "srv", JoinKind.LeftOuter),
  WithServer = Table.ExpandTableColumn(Joined, "srv", {"server_wire_roll_cert_number"}, {"server_wire_roll_cert_number"}),

  // 4) Keep rows where user changed the value and it's nonblank
  WithChangedFlag = Table.AddColumn(
    WithServer, "changed",
    each
      let
        userVal   = Record.FieldOrDefault(_, "wire_roll_cert_number", null),
        serverVal = Record.FieldOrDefault(_, "server_wire_roll_cert_number", null)
      in
        userVal <> null
        and Text.Trim(Text.From(userVal)) <> ""
        and userVal <> serverVal,
    type logical
  ),
  ChangesOnly = Table.SelectRows(WithChangedFlag, each [changed] = true),

  // 5) Final payload for PATCH
  WireSetCertChanges = Table.SelectColumns(ChangesOnly, {"asset_id","wire_roll_cert_number"}, MissingField.UseNull),

  // Single source of truth for the output schema
  ResultType = type table [
      ok = logical,
      status = number,
      message = text,
      asset_id = Int64.Type,
      new_wire_roll_cert_number = text,
      attempted_wire_roll_cert_number = text
  ],

  // Empty result with headers & types (prevents Excel's "no columns" warning)
  EmptyResult = Table.FromRows({}, ResultType),

  Results =
    if Table.RowCount(WireSetCertChanges) = 0 then
      EmptyResult
    else
      let
        Records     = Table.ToRecords(WireSetCertChanges),
        ResultsList = List.Transform(Records, each PostWireSetCertUpdate(_))
      in
        Table.FromRecords(ResultsList, ResultType)
in
  Results