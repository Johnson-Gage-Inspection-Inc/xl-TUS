VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ThisWorkbook.Connections("Query - Survey").Refresh
    ThisWorkbook.Sheets("Survey").Calculate
    ThisWorkbook.Connections("Query - Comparison").Refresh
End Sub

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
        ' Check user setting to see if warnings should be hidden
        If Not UserSettings.GetHideHeaderMismatchWarning() Then
            Dim fullMessage As String
            fullMessage = label & ": Header mismatch detected. Using existing headers." & vbCrLf & vbCrLf & "Mapping (Worksheet) <- (Clipboard):" & vbCrLf & msg
            Dim userResponse As Boolean
            userResponse = HeaderMismatchDialog.ShowHeaderMismatchWarning("Header Diff for " & label, fullMessage)
            If Not userResponse Then Exit Sub
        Else
            ' Still log the information to immediate window for debugging
            Debug.Print "Header mismatch for " & label & " (warning suppressed by user setting)"
            Debug.Print "Mapping (Worksheet) <- (Clipboard):"
            Debug.Print msg
        End If
    End If

    ' Check if row 2 is a blank spacer row (optional)
    Dim hasSpacerRow As Boolean
    hasSpacerRow = False
    fields = Split(Replace(cleanLines(2), vbTab, ","), ",")
    Dim isBlankRow As Boolean
    isBlankRow = True
    For j = LBound(fields) To UBound(fields)
        If Trim(fields(j)) <> "" Then
            isBlankRow = False
            Exit For
        End If
    Next j
    hasSpacerRow = isBlankRow

    ' Validate all data rows (start from row 3 if spacer exists, row 2 if not)
    Dim dataStartIndex As Long
    dataStartIndex = IIf(hasSpacerRow, 3, 2)
    
    For i = dataStartIndex To cleanLines.count
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

    Dim numRows As Long: numRows = cleanLines.count - (IIf(hasSpacerRow, 2, 1))
    Dim numCols As Long: numCols = colCount
    Dim data() As Variant: ReDim data(1 To numRows, 1 To numCols)

    For i = 1 To numRows
        fields = Split(Replace(cleanLines(i + (IIf(hasSpacerRow, 2, 1))), vbTab, ","), ",")
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
Sub NegateSelectedValuesIfInDaqBookRange()
    Dim validRange As Range
    Dim intersectRange As Range
    Dim area As Range
    Dim values As Variant
    Dim i As Long, j As Long
    Dim numRows As Long, numCols As Long

    ' Use three specific ranges: B3:O, R3:AE, AH3:AS
    Dim range1 As Range, range2 As Range, range3 As Range
    Set range1 = Me.Range("B3:O" & Me.Rows.count)
    Set range2 = Me.Range("R3:AE" & Me.Rows.count)
    Set range3 = Me.Range("AH3:AS" & Me.Rows.count)
    Set validRange = Application.Union(range1, range2, range3)

    On Error Resume Next
    Set intersectRange = Application.Intersect(Selection, validRange)
    On Error GoTo 0

    If intersectRange Is Nothing Then
        MsgBox "No selected cells are within the valid DaqBook data ranges (B3:O, R3:AE, AH3:AS).", vbInformation
        Exit Sub
    End If

    ' Turn off screen updating and calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Process each area in the intersection (handles non-contiguous selections)
    For Each area In intersectRange.Areas
        numRows = area.Rows.count
        numCols = area.Columns.count
        
        ' Read values into array for faster processing
        If numRows = 1 And numCols = 1 Then
            ' Single cell - handle directly
            If IsNumeric(area.value) And Not IsEmpty(area.value) Then
                area.value = -CDbl(area.value)
            End If
        Else
            ' Multiple cells - use array operations
            values = area.value
            
            ' Process the array
            If numRows = 1 Then
                ' Single row
                For j = 1 To numCols
                    If IsNumeric(values(1, j)) And Not IsEmpty(values(1, j)) Then
                        values(1, j) = -CDbl(values(1, j))
                    End If
                Next j
            ElseIf numCols = 1 Then
                ' Single column
                For i = 1 To numRows
                    If IsNumeric(values(i, 1)) And Not IsEmpty(values(i, 1)) Then
                        values(i, 1) = -CDbl(values(i, 1))
                    End If
                Next i
            Else
                ' Multiple rows and columns
                For i = 1 To numRows
                    For j = 1 To numCols
                        If IsNumeric(values(i, j)) And Not IsEmpty(values(i, j)) Then
                            values(i, j) = -CDbl(values(i, j))
                        End If
                    Next j
                Next i
            End If
            
            ' Write the modified array back to the range
            area.value = values
        End If
    Next area
    
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
