let
    Source = try Table.SelectRows(customers, each [CompanyId] = ClientCompanyId) otherwise customers
in
    Source