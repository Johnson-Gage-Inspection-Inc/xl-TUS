VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    ThisWorkbook.Sheets("Comparison_Report").Calculate
    ThisWorkbook.Sheets("Main").Calculate
    ThisWorkbook.Sheets("CERT").Calculate
End Sub
Private Sub Worksheet_Calculate()
    ThisWorkbook.Sheets("Main").Calculate
    ThisWorkbook.Sheets("CERT").Calculate
End Sub
