let
    TokenText = Text.Trim(Text.FromBinary(
        Web.Contents("https://jgiquality.sharepoint.com/sites/JGI/Shared%20Documents/General/apikey.txt")
    ))
in
    TokenText