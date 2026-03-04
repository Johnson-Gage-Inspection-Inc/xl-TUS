Attribute VB_Name = "UnitFormatting"
Option Explicit

' ---------------------------------------------------------------------------
' UnitFormatting.bas
'
' Applies temperature-unit number formats (°F or °C) to all cells that use
' degree-symbol format codes. Called from Workbook_Open and whenever the
' Unit named range changes value.
'
' The workbook's "Unit" named range returns "°F" or "°C".
' The "_unit" named range returns just "F" or "C" (no degree symbol).
'
' Format codes affected (7 unique patterns):
'   0°F          ?  0°C
'   0.0°F        ?  0.0°C
'   0.0 °F       ?  0.0 °C
'   #°F          ?  #°C
'   ±0°F         ?  ±0°C
'   ±0.0°F       ?  ±0.0°C
'   [>0]+0.0°F;[<0]-0.0°F; 0.0°F  ?  [>0]+0.0°C;[<0]-0.0°C; 0.0°C
' ---------------------------------------------------------------------------

' Cached last-applied unit letter so we can detect changes
Public LastAppliedUnit As String

Public Sub ApplyUnitFormats()
    Dim deg As String:      deg = ChrW$(176)            ' °
    
    ' Read the current unit from the named range
    Dim unitStr As String
    On Error Resume Next
    unitStr = Range("Unit").value   ' "°F" or "°C"
    On Error GoTo 0
    If unitStr = "" Then unitStr = deg & "F"

    Dim unitLetter As String:   unitLetter = Right$(unitStr, 1)   ' "F" or "C"

    ' Skip if formats already match
    If LastAppliedUnit = unitLetter Then Exit Sub
    LastAppliedUnit = unitLetter

    ' Determine target and source substrings for replacement
    Dim targetDeg As String     ' what we want in format codes
    Dim sourceDeg As String     ' what we want to replace
    targetDeg = deg & unitLetter
    sourceDeg = deg & IIf(unitLetter = "F", "C", "F")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Iterate only sheets known to have degree-formatted cells.
    ' If a sheet doesn't exist (e.g., in a trimmed workbook), skip it.
    Dim sheetNames As Variant
    sheetNames = Array("CMCs", "Main", "CERT", "Data_Sheet", "Data_Sheet_15_28", _
                       "Data_Sheet_29_40", "Comparison_Report", "TUS_Worksheet", _
                       "Interp")

    Dim i As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim fmt As String

    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
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

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

