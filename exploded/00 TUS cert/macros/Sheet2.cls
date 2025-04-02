VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Range("D19, G19, J19, M19, D31, G31, J31, M31")) Is Nothing Then Exit Sub

    On Error GoTo HandleError
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim iCells As Variant, rRanges As Variant
    
        iCells = Array("D19", "G19", "J19", "M19", "D31", "G31", "J31", "M31")
    rRanges = Array("D10:D18", "G10:G18", "J10:J18", "M10:M18", _
                    "D22:D30", "G22:G30", "J22:J30", "M22:M30")
    
    Dim i As Long
    For i = 0 To UBound(iCells)
        With Range(rRanges(idx))
            .NumberFormat = IIf(Range(iCells(i)).Value = 0, "0°F", "0.0°F")
        End With
    Next i

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

