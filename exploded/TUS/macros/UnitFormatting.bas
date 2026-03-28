Attribute VB_Name = "UnitFormatting"
Option Explicit

' ---------------------------------------------------------------------------
' UnitFormatting.bas
'
' Swaps temperature-unit number formats between °F and °C to match the
' workbook's "Unit" named range ("°F" or "°C").
'
' Triggered by Worksheet_Calculate on UUTRangeTol (Sheet17.cls)
'
' Format codes affected (5 unique patterns found by find_unit_formats.py):
'   0\°\F                                          ->  0\°\C
'   0.0\°\F                                        ->  0.0\°\C
'   0.0\ \°\F                                      ->  0.0\ \°\C
'   \±0.0\°\F                                      ->  \±0.0\°\C
'   [>0]\+0.0\°\F;[<0]\-0.0\°\F;\ 0.0\°\F        ->  [>0]\+0.0\°\C;...
'
' All patterns contain the literal substring °F (or °C), so a single
' Replace() call handles every variation.
'
' Current formatting state is detected from Main!D15.NumberFormat rather
' than a module-level variable, so it survives VBA project resets.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' ApplyUnitFormats
'
' Reads the "Unit" named range and replaces °F<->°C in the NumberFormat
' property of every cell that carries a degree-symbol format code.
'
' For large data sheets (Data_Sheet*) where only a single cell (J6) uses
' the degree format, we target that cell directly instead of iterating
' ~14 000 cells per sheet.
'
' The current formatting state is detected from Main!D15.NumberFormat
' (a cell with format 0\?\F or 0\?\C).  This avoids a module-level
' variable that would reset on VBA project reset / unhandled error.
' ---------------------------------------------------------------------------
Public Sub ApplyUnitFormats()
    Dim deg As String:  deg = ChrW$(176)  ' °

    ' Read the current unit from the named range
    Dim unitStr As String
    On Error Resume Next
    unitStr = Range("Unit").value
    On Error GoTo 0
    If unitStr = "" Then unitStr = deg & "F"

    Dim unitLetter As String: unitLetter = Right$(unitStr, 1)  ' "F" or "C"

    ' Detect current formatting state from a known cell rather than a
    ' module-level variable (survives VBA project resets)
    Dim probeFmt As String
    probeFmt = ThisWorkbook.Sheets("Main").Range("D15").NumberFormat
    If InStr(1, probeFmt, deg & unitLetter, vbBinaryCompare) > 0 Then
        Exit Sub  ' formats already match the target unit
    End If

    ' Determine source -> target substrings for replacement
    Dim targetDeg As String
    Dim sourceDeg As String
    targetDeg = deg & unitLetter
    sourceDeg = deg & IIf(unitLetter = "F", "C", "F")

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- Large sheets: target only the known cell (J6) ------------------
    Dim bigSheets As Variant
    bigSheets = Array("Data_Sheet", "Data_Sheet_15_28", "Data_Sheet_29_40")

    Dim ws As Worksheet
    Dim i As Long
    Dim fmt As String

    For i = LBound(bigSheets) To UBound(bigSheets)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(bigSheets(i))
        On Error GoTo Cleanup
        If Not ws Is Nothing Then
            fmt = ws.Range("J6").NumberFormat
            If InStr(1, fmt, sourceDeg, vbBinaryCompare) > 0 Then
                ws.Range("J6").NumberFormat = Replace(fmt, sourceDeg, targetDeg)
            End If
        End If
    Next i

    ' ---- Smaller sheets: iterate UsedRange ------------------------------
    Dim smallSheets As Variant
    smallSheets = Array("Main", "CERT", "Comparison_Report", _
                        "TUS_Worksheet", "Interp")

    Dim cell As Range

    For i = LBound(smallSheets) To UBound(smallSheets)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(smallSheets(i))
        On Error GoTo Cleanup
        If ws Is Nothing Then GoTo NextSheet

        For Each cell In ws.UsedRange
            fmt = cell.NumberFormat
            If InStr(1, fmt, sourceDeg, vbBinaryCompare) > 0 Then
                cell.NumberFormat = Replace(fmt, sourceDeg, targetDeg)
            End If
        Next cell

        Set ws = Nothing
NextSheet:
        Set ws = Nothing
    Next i

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
