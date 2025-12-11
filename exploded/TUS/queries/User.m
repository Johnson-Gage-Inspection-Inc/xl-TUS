let
    required = {"User"},
    empty    = #table(required, {}),

    // --- Try API ---
    fromApi =
        try
            let
                resp  = Web.Contents(
                    "https://jgiapi.com",
                    [ RelativePath = "whoami",
                      ManualStatusHandling = {400,401,403,404,500,502,503,504} ]
                ),
                js    = Json.Document(resp),
                email = Record.FieldOrDefault(js, "user", null)
            in
                if email <> null and email <> "" then #table(required, {{email}}) else empty
        otherwise empty,

    // --- Fallback: prior worksheet contents (soft cache) ---
    fromCache =
        try Excel.CurrentWorkbook(){[Name="User"]}[Content] otherwise empty,

    chosen = if Table.RowCount(fromApi) > 0 then fromApi else fromCache,

    // enforce column and order
    final  = Table.SelectColumns(chosen, required, MissingField.UseNull)
in
    final