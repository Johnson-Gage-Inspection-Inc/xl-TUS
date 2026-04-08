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

Private Sub Worksheet_Calculate()
    On Error GoTo HandleError
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim iCells As Variant, rRanges As Variant
    Dim unitStr As String, unitLetter As String, deg As String
    
    iCells = Array("D19", "G19", "J19", "M19", "D31", "G31", "J31", "M31")
    rRanges = Array("D10:D18", "G10:G18", "J10:J18", "M10:M18", _
                    "D22:D30", "G22:G30", "J22:J30", "M22:M30")

    deg = ChrW$(176)
    On Error Resume Next
    unitStr = CStr(ThisWorkbook.names("Unit").RefersToRange.value)
    If unitStr = "" Then unitStr = CStr(Application.Range("Unit").value)
    On Error GoTo HandleError
    If unitStr = "" Then unitStr = deg & "F"
    unitLetter = Right$(unitStr, 1)
    If unitLetter <> "F" And unitLetter <> "C" Then unitLetter = "F"
    
    Dim i As Long
    For i = 0 To UBound(iCells)
        With Range(rRanges(i))
            .NumberFormat = IIf(Range(iCells(i)).value = 0, _
                                "0" & deg & unitLetter, _
                                "0.0" & deg & unitLetter)
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
