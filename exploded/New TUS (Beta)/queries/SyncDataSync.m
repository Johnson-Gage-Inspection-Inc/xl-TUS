let
  respTry =
    try Web.Contents(
      "https://jgiapi.com",
      [
        RelativePath = "data-sync/",
        Query = [ init = "false" ],
        ManualStatusHandling = {200,202,204,400,401,403,404,409,422,500},
        Timeout = #duration(0,0,60,0)
      ]
    ),

  result =
    if respTry[HasError] then
      #table(
        {"ok", "status", "message"},
        {
          { false, null, "Client-side request error (see data source/privacy settings)" }
        }
      )
    else
      let
        resp   = respTry[Value],
        status = try Value.Metadata(resp)[Response.Status] otherwise null,
        body   = try Json.Document(resp) otherwise null
      in
        #table(
          {"ok", "status", "message"},
          {
            {
              List.Contains({200,202,204}, status),
              status,
              if body <> null and Record.HasFields(body, "status")
              then "Data sync: " & Text.From(body[status])
              else "Triggered"
            }
          }
        )
in
  result