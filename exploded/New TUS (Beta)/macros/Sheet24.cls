VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Calculate
    ThisWorkbook.Sheets("Main").Calculate
End Sub
Private Sub Worksheet_Calculate()
    ' Force full recalculation by making cells dirty
    With ThisWorkbook.Sheets("Main")
        Application.Calculation = xlCalculationManual
        .Range("P5:P44").Dirty
        Application.Calculation = xlCalculationAutomatic
    End With
    
    ThisWorkbook.Sheets("CERT").Calculate
End Sub
