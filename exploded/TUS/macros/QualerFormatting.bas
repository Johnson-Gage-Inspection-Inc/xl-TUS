Attribute VB_Name = "QualerFormatting"
Public Function GetFormattedWorkItemNumber(OrderInput As String, Optional ItemInput As String = "") As String
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
    Dim mainNum As Long
    If Not TryParseLong(mainRaw, mainNum) Then
        GetFormattedWorkItemNumber = vbNullString
        Exit Function
    End If

    Dim mainOrder As String
    mainOrder = Format(mainNum, "000000")

    Dim childOrder As String
    If UBound(parts) >= 1 Then
        Dim childNum As Long
        If Not TryParseLong(parts(1), childNum) Then
            GetFormattedWorkItemNumber = vbNullString
            Exit Function
        End If
        childOrder = "." & Format(childNum, "00")
    Else
        childOrder = ""
    End If

    If Len(Trim$(ItemInput)) = 0 Then
        GetFormattedWorkItemNumber = prefix & mainOrder & childOrder
        Exit Function
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
    Dim itemNum As Long
    If Not TryParseLong(itemMain, itemNum) Then
        GetFormattedWorkItemNumber = vbNullString
        Exit Function
    End If
    paddedItem = Format(itemNum, "00") & itemRev

    ' Return final formatted string
    GetFormattedWorkItemNumber = prefix & mainOrder & childOrder & "-" & paddedItem
End Function

Private Function TryParseLong(ByVal textValue As String, ByRef parsedValue As Long) As Boolean
    Dim cleaned As String
    cleaned = Trim$(textValue)

    If Len(cleaned) = 0 Then Exit Function
    If cleaned Like "*[!0-9]*" Then Exit Function

    On Error GoTo ParseFailed
    parsedValue = CLng(cleaned)
    TryParseLong = True
    Exit Function

ParseFailed:
    TryParseLong = False
End Function
