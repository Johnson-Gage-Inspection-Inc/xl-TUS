let
    Source = Employees,
    response = Web.Contents("https://api.jgiquality.com", [RelativePath = "whoami"]),
    json = Json.Document(response),
    email = json[user],
    #"Filtered Rows" = Table.SelectRows(Source, each ([login_email] = email))
in
    #"Filtered Rows"