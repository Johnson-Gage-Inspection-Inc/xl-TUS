Attribute VB_Name = "QualerFormatting"
Public Function GetFormattedWorkItemNumber(OrderInput As String, ItemInput As String) As String
    Dim prefix As String: prefix = "56561-"
    Dim rawOrder As String

    ' Strip prefix if present
    If Left(OrderInput, Len(prefix)) = prefix Then
        rawOrder = Mid(OrderInput, Len(prefix) + 1)
    Else
        rawOrder = OrderInput
    End If

    ' Split into main and optional child
    Dim parts() As String
    parts = Split(rawOrder, ".")

    Dim mainRaw As String: mainRaw = parts(0)
    Dim mainOrder As String
    mainOrder = Format(CLng(mainRaw), "000000")

    Dim childOrder As String
    If UBound(parts) >= 1 Then
        childOrder = "." & Format(CLng(parts(1)), "00")
    Else
        childOrder = ""
    End If

    ' Parse ItemInput
    Dim itemMain As String, itemRev As String
    Dim rPos As Long: rPos = InStr(1, ItemInput, "R", vbTextCompare)
    If rPos > 0 Then
        itemMain = Left(ItemInput, rPos - 1)
        itemRev = Mid(ItemInput, rPos)
    Else
        itemMain = ItemInput
        itemRev = ""
    End If

    Dim paddedItem As String
    paddedItem = Format(CLng(itemMain), "00") & itemRev

    ' Return final formatted string
    GetFormattedWorkItemNumber = prefix & mainOrder & childOrder & "-" & paddedItem
End Function

