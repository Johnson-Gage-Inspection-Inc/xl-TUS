let
    // Read wirelots from Excel named range
    RawWirelots = Excel.CurrentWorkbook(){[Name="Wirelots"]}[Content][Column1],
    WirelotList = List.Select(List.Distinct(List.Transform(RawWirelots, each Text.Upper(Text.Trim(_)))), each _ <> ""),

    // Generate filename candidates
    GetCharCodes = List.Transform(WirelotList, each {Text.Start(_, 6), Text.End(_, 1)}),
    FileCandidates = List.Combine(
        List.Transform(GetCharCodes, each
            let
                number = _{0},
                letter = _{1},
                code = Character.ToNumber(letter),
                prev = if code > 65 then Character.FromNumber(code - 1) else letter,
                next = if code < 90 then Character.FromNumber(code + 1) else letter,
                paths = {
                    number & letter & ".xls",
                    number & prev & "-" & letter & ".xls",
                    number & letter & "-" & next & ".xls"
                }
            in paths)
    ),
    DistinctFiles = List.Distinct(FileCandidates),

    // Build full SharePoint URLs
    BaseUrl = "https://jgiquality.sharepoint.com/sites/JGI/Shared%20Documents/Pyro/Pyro_Standards/",
    FileUrls = List.Transform(DistinctFiles, each [File=_ , Url=BaseUrl & _]),

    // Try loading each file and extracting two blocks of data
    LoadFile = (entry) =>
        let
            source = try Excel.Workbook(Web.Contents(entry[Url]), null, true) otherwise null,
            sheet = if source <> null then try source{[Item="TC Form", Kind="Sheet"]}[Data] otherwise null else null,
            result =
                if sheet <> null then
                    let
                        Wirelot = try Record.Field(sheet{650}, "Column2") otherwise null,
                        Row1 = try List.Transform({"Column11".."Column15"}, each Record.Field(sheet{652}, _)) otherwise {},
                        Row2 = try List.Transform({"Column11".."Column15"}, each Record.Field(sheet{659}, _)) otherwise {},
                        SecondLot = try Record.Field(sheet{690}, "Column2") otherwise null,
                        Row3 = try List.Transform({"Column11".."Column15"}, each Record.Field(sheet{692}, _)) otherwise {},
                        Row4 = try List.Transform({"Column11".."Column15"}, each Record.Field(sheet{699}, _)) otherwise {},
                        Records = {
                            [Wirelot=Wirelot, Block="Top", Col1=Row1{0}, Col2=Row1{1}, Col3=Row1{2}, Col4=Row1{3}, Col5=Row1{4}],
                            [Wirelot=Wirelot, Block="Bottom", Col1=Row2{0}, Col2=Row2{1}, Col3=Row2{2}, Col4=Row2{3}, Col5=Row2{4}],
                            [Wirelot=SecondLot, Block="Top", Col1=Row3{0}, Col2=Row3{1}, Col3=Row3{2}, Col4=Row3{3}, Col5=Row3{4}],
                            [Wirelot=SecondLot, Block="Bottom", Col1=Row4{0}, Col2=Row4{1}, Col3=Row4{2}, Col4=Row4{3}, Col5=Row4{4}]
                        }
                    in
                        List.Select(Records, each _[Wirelot] <> null)
                else {}
        in
            result,

    AllResults = List.Combine(List.Transform(FileUrls, LoadFile)),
    Output = Table.FromRecords(AllResults)
in
    Output