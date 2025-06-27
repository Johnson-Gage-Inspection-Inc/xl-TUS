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

'=======================
' Core reusable logic
'=======================
Private Sub PasteChannelBlock(sheetName As String, startCellAddress As String, channelCount As Long, label As String, startingChannel As Long)
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

    ' Build expected headers: Time, 1...N
    ReDim expectedHeader(0 To channelCount)
    expectedHeader(0) = "Time"
    For i = 1 To channelCount
        expectedHeader(i) = CStr(startingChannel + i - 1)
    Next i
    colCount = channelCount + 1

    ' Get clipboard text
    On Error GoTo ClipboardError
    Set clipboard = New MSForms.DataObject
    clipboard.GetFromClipboard
    rawText = clipboard.GetText
    On Error GoTo 0

    ' Split by lines
    lines = Split(rawText, vbCrLf)
    If UBound(lines) < 2 Then
        MsgBox "Clipboard content is too short. Expected at least 3 lines.", vbCritical
        Exit Sub
    End If

    ' Clean empty lines
    Dim cleanLines As Collection
    Set cleanLines = New Collection
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) <> "" Then cleanLines.Add lines(i)
    Next i
    If cleanLines.count < 3 Then
        MsgBox "No data rows found after headers.", vbCritical
        Exit Sub
    End If

    ' Parse header
    pastedHeader = Split(Replace(cleanLines(1), vbTab, ","), ",")
    If UBound(pastedHeader) + 1 > colCount Then
        MsgBox "Header column count mismatch for " & label & ". Expected " & colCount & " columns.", vbCritical
        Exit Sub
    End If
    
    Dim msg As String
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


    ' Validate row 2 is blank
    fields = Split(Replace(cleanLines(2), vbTab, ","), ",")
    Dim isEmptySpacer As Boolean
    isEmptySpacer = True
    For j = LBound(fields) To UBound(fields)
        If Trim(fields(j)) <> "" Then
            isEmptySpacer = False
            Exit For
        End If
    Next j
    If Not isEmptySpacer Then
        MsgBox "Expected a blank spacer row after headers.", vbCritical
        Exit Sub
    End If


    ' Validate data rows
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

    ' Get target range
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set startCell = ws.Range(startCellAddress)
    rowOffset = 1
    
    ' Build and paste all values using a 2D array (faster)
    Dim data() As Variant
    Dim numRows As Long: numRows = cleanLines.count - 2 ' skip header + blank row
    Dim numCols As Long: numCols = colCount
    ReDim data(1 To numRows, 1 To numCols)
    
    For i = 1 To numRows
        fields = Split(Replace(cleanLines(i + 2), vbTab, ","), ",")
        For j = 0 To UBound(fields)
            data(i, j + 1) = Trim(fields(j))
        Next j
    Next i
    
    ' Write to sheet in one go
    startCell.Offset(1, 0).Resize(numRows, numCols).Value = data
    Exit Sub

ClipboardError:
    MsgBox "Could not read clipboard. Ensure you've copied the TSV data.", vbCritical
End Sub

