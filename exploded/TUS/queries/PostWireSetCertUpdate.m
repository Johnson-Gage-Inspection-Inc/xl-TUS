// Connection Properties (Usage tab)
//   BackgroundQuery:       True
//   RefreshOnFileOpen:     False
//   RefreshPeriod:         0
//   RefreshWithRefreshAll: True
//   EnableFastDataLoad:    False

// PostWireSetCertUpdate (GET + query params; works with Org auth in Excel Web)
let
  PostWireSetCertUpdate = (row as record) as record =>
    let
      urlBase = "https://jgiapi.com",
      path    = "wire-set-certs/update",

      // Build query params dynamically - only include non-null values
      buildQueryParams = (row as record) as record =>
        let
          baseParams = [asset_id = Text.From(Record.Field(row, "asset_id"))],
          withWireRoll = 
            if Record.HasFields(row, "wire_roll_cert_number") and Record.Field(row, "wire_roll_cert_number") <> null 
            then Record.AddField(baseParams, "wire_roll_cert_number", Text.From(Record.Field(row, "wire_roll_cert_number")))
            else baseParams,
          withTraceability = 
            if Record.HasFields(row, "traceability_number") and Record.Field(row, "traceability_number") <> null 
            then Record.AddField(withWireRoll, "traceability_number", Text.From(Record.Field(row, "traceability_number")))
            else withWireRoll
        in
          withTraceability,

      queryParams = buildQueryParams(row),

      // Send query params (no Content=) so Excel Web will use Org auth
      responseTry = try Web.Contents(
        urlBase,
        [
          RelativePath = path,
          Query = queryParams,
          ManualStatusHandling = {200, 204, 400, 401, 403, 404, 409, 422, 500},
          Timeout = #duration(0, 0, 30, 0)
        ]
      ),

      result =
        if responseTry[HasError] then
          let
            e       = responseTry[Error],
            errText = try Text.FromBinary(Json.FromValue(e)) otherwise null
          in
            [
              ok = false,
              status = null,
              message = "Client-side request error",
              detail = errText,
              asset_id = Record.Field(row, "asset_id"),
              attempted_wire_roll_cert_number = try Record.Field(row, "wire_roll_cert_number") otherwise null,
              attempted_traceability_number = try Record.Field(row, "traceability_number") otherwise null
            ]
        else
          let
            resp   = responseTry[Value],
            status = try Value.Metadata(resp)[Response.Status] otherwise null,
            body   = try Json.Document(resp) otherwise null,
            msg =
              if status = 200 or status = 204 then "Updated"
              else if body <> null and Record.HasFields(body, "detail") then body[detail]
              else if body <> null and Record.HasFields(body, "message") then body[message]
              else "Update failed"
          in
            if status = 200 or status = 204 then
              [
                ok = true,
                status = status,
                message = msg,
                asset_id = Record.Field(row, "asset_id"),
                new_wire_roll_cert_number = try Record.Field(row, "wire_roll_cert_number") otherwise null,
                new_traceability_number = try Record.Field(row, "traceability_number") otherwise null
              ]
            else
              [
                ok = false,
                status = status,
                message = msg,
                asset_id = Record.Field(row, "asset_id"),
                attempted_wire_roll_cert_number = try Record.Field(row, "wire_roll_cert_number") otherwise null,
                attempted_traceability_number = try Record.Field(row, "traceability_number") otherwise null
              ]
    in
      result
in
  PostWireSetCertUpdate