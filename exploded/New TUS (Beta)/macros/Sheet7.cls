VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=======================
' Entry points (one per button)
'=======================
Public Sub PasteChannels1to14()
    PasteChannelBlock "DaqBook_RAW_Data", "A2", 14, "DataForChannels1to14", 1
End Sub

Public Sub PasteChannels15to28()
    PasteChannelBlock "DaqBook_RAW_Data", "Q2", 14, "DataForChannels15to28", 15
End Sub

Public Sub PasteChannels29to40()
    PasteChannelBlock "DaqBook_RAW_Data", "AG2", 12, "DataForChannels29to40", 29
End Sub

Public Sub TruncateChannels1to14()
    Call TruncateTableByName("DataForChannels1to14")
End Sub

Public Sub TruncateChannels15to28()
    Call TruncateTableByName("DataForChannels15to28")
End Sub

Public Sub TruncateChannels29to40()
    Call TruncateTableByName("DataForChannels29to40")
End Sub

'=======================
' Core reusable logic
'=======================
Public Sub PasteChannelBlock(sheetName As String, startCellAddress As String, channelCount As Long, label As String, startingChannel As Long, Optional tsvTextSource As String = "")
    Dim clipboard As MSForms.DataObject
    Dim rawText As String
    Dim lines As Variant, fields As Variant
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim startCell As Range
    Dim expectedHeader() As String
    Dim pastedHeader() As String
    Dim headerMismatch As Boolean
    Dim rowOffset As Long
    Dim colCount As Long
    Dim msg As String

    ' Build expected headers: Time, 1...N
    ReDim expectedHeader(0 To channelCount)
    expectedHeader(0) = "Time"
    For i = 1 To channelCount
        expectedHeader(i) = CStr(startingChannel + i - 1)
    Next i
    colCount = channelCount + 1

    ' Get TSV data: from clipboard or argument
    If Len(tsvTextSource) = 0 Then
        On Error GoTo ClipboardError
        Set clipboard = New MSForms.DataObject
        clipboard.GetFromClipboard
        rawText = clipboard.GetText
        On Error GoTo 0
    Else
        rawText = tsvTextSource
    End If

    Call TruncateTableByName(label)

    ' Split into lines and clean
    lines = Split(rawText, vbCrLf)
    Dim cleanLines As Collection: Set cleanLines = New Collection
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then cleanLines.Add lines(i)
    Next i
    If cleanLines.count < 3 Then
        MsgBox "Not enough non-empty lines in data. Expected at least header + spacer + one row.", vbCritical
        Exit Sub
    End If

    ' Parse header
    pastedHeader = Split(Replace(cleanLines(1), vbTab, ","), ",")
    If UBound(pastedHeader) + 1 > colCount Then
        MsgBox "Header column count mismatch for " & label & ". Expected " & colCount & " columns.", vbCritical
        Exit Sub
    End If

    headerMismatch = False
    For j = 0 To Application.Min(UBound(expectedHeader), UBound(pastedHeader))
        If j = 0 Then
            If LCase(Trim(pastedHeader(0))) <> "time" Then
                headerMismatch = True
                msg = msg & "Time <- " & pastedHeader(0) & vbCrLf
            End If
        ElseIf val(pastedHeader(j)) <> val(expectedHeader(j)) Then
            headerMismatch = True
            msg = msg & expectedHeader(j) & " <- " & pastedHeader(j) & vbCrLf
        End If
    Next j
    If headerMismatch Then
        MsgBox label & ": Header mismatch detected. Using existing headers. Mapping:" & vbCrLf & msg, vbExclamation, "Header Diff for " & label
    End If

    ' Validate spacer row
    fields = Split(Replace(cleanLines(2), vbTab, ","), ",")
    For j = LBound(fields) To UBound(fields)
        If Trim(fields(j)) <> "" Then
            MsgBox "Expected a blank spacer row after headers.", vbCritical
            Exit Sub
        End If
    Next j

    ' Validate all data rows
    For i = 3 To cleanLines.count
        fields = Split(Replace(cleanLines(i), vbTab, ","), ",")
        If UBound(fields) + 1 < 2 Then
            MsgBox "Row " & i & " has too few columns (must include Time and at least one channel).", vbCritical
            Exit Sub
        End If
        If UBound(fields) + 1 > colCount Then
            MsgBox "Row " & i & " has too many columns. Expected up to " & colCount & ".", vbCritical
            Exit Sub
        End If
        If Not IsDate(Trim(fields(0))) Then
            MsgBox "Row " & i & " has invalid time value: " & fields(0), vbCritical
            Exit Sub
        End If
        For j = 1 To UBound(fields)
            If Not IsNumeric(fields(j)) Then
                MsgBox "Row " & i & ", column " & j & " has non-numeric data: " & fields(j), vbCritical
                Exit Sub
            End If
        Next j
    Next i

    ' Paste data
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set startCell = ws.Range(startCellAddress)

    Dim numRows As Long: numRows = cleanLines.count - 2
    Dim numCols As Long: numCols = colCount
    Dim data() As Variant: ReDim data(1 To numRows, 1 To numCols)

    For i = 1 To numRows
        fields = Split(Replace(cleanLines(i + 2), vbTab, ","), ",")
        For j = 0 To UBound(fields)
            data(i, j + 1) = Trim(fields(j))
        Next j
    Next i

    startCell.Offset(1, 0).Resize(numRows, numCols).value = data
    Exit Sub

ClipboardError:
    MsgBox "Could not read clipboard. Ensure you've copied the TSV data.", vbCritical
End Sub
Public Sub TruncateTableByName(tableName As String)
    Dim tbl As ListObject
    Dim ws As Worksheet
    
    ' Loop through all sheets to find the table
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not tbl Is Nothing Then Exit For
    Next ws

    If tbl Is Nothing Then
        MsgBox "Table '" & tableName & "' not found in this workbook.", vbCritical
        Exit Sub
    End If

    If tbl.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    tbl.DataBodyRange.Rows.Delete
End Sub
