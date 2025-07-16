VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Sub UpdateWorkItemDetails()
    ' TruncateWorkItemDetailsTable
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ThisWorkbook.Connections("Query - WorkItemDetails").Refresh
    Me.Calculate
End Sub
Private Sub TruncateWorkItemDetailsTable()
    Dim lo As ListObject
    Set lo = Me.ListObjects("WorkItemDetails")
    If lo.DataBodyRange Is Nothing Then Exit Sub ' Already empty
    lo.DataBodyRange.Rows.Delete
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Calculate
End Sub
Private Sub Worksheet_Calculate()
    ThisWorkbook.Sheets("Header_Info").Calculate
    ThisWorkbook.Sheets("Customer Table").Calculate
End Sub
