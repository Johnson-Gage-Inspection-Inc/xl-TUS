let
    response = Web.Contents("https://jgiapi.com", [RelativePath = "employees"]),
    json = Json.Document(response),
    #"Converted to Table" = Table.FromList(json, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Expanded Column1" = Table.ExpandRecordColumn(#"Converted to Table", "Column1", {"alias", "company_id", "culture_name", "culture_ui_name", "departments", "employee_id", "first_name", "image_url", "is_deleted", "is_locked", "last_name", "last_seen_date_utc", "login_email", "office_phone", "subscription_email", "subscription_phone", "title"}, {"alias", "company_id", "culture_name", "culture_ui_name", "departments", "employee_id", "first_name", "image_url", "is_deleted", "is_locked", "last_name", "last_seen_date_utc", "login_email", "office_phone", "subscription_email", "subscription_phone", "title"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Column1", each [is_deleted] = false),
    #"Merged Columns" = Table.CombineColumns(#"Filtered Rows",{"first_name", "last_name"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Name"),
    #"Expanded departments" = Table.ExpandListColumn(#"Merged Columns", "departments"),
    #"Filtered Rows1" = Table.SelectRows(#"Expanded departments", each true),
    #"Expanded departments1" = Table.ExpandRecordColumn(#"Filtered Rows1", "departments", {"name"}, {"departments.name"}),
    // Tag each row with whether it's the Pyrometry Department
    #"Added Pyro Flag" = Table.AddColumn(#"Expanded departments1", "IsPyro", each [departments.name] = "Pyrometry Department", type logical),

    // Group by employee, keep one row per employee
    #"Grouped Rows" = Table.Group(#"Added Pyro Flag", {"employee_id", "Name", "login_email"}, {
        {"IsPyro", each List.AnyTrue([IsPyro]), type logical}
    }),

    // Optional: sort if needed
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Name", Order.Ascending}})

in
    #"Sorted Rows"