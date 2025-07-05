VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub UpdateSurvey()
    ' TruncateWorkItemDetailsTable
    ThisWorkbook.Connections("Query - Survey").Refresh
End Sub
Private Sub TruncateSurveyTable()
    Dim lo As ListObject
    Set lo = Me.ListObjects("Survey")
    If lo.DataBodyRange Is Nothing Then Exit Sub ' Already empty
    lo.DataBodyRange.Rows.Delete
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Calculate
    ThisWorkbook.Sheets("Main").Calculate
End Sub
Private Sub Worksheet_Calculate()
    ThisWorkbook.Sheets("Main").Range("P5:P44").Calculate
    ThisWorkbook.Sheets("CERT").Calculate
    Sheet6.ForceRecalc_Alerts
End Sub
