VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Sub UpdateClientAssetAttributes()
    TruncateClientAssetAttributesTable
    ThisWorkbook.Connections("Query - ClientAssetAttributes_GetAssetAttributesByAssetid").Refresh
End Sub
Private Sub TruncateClientAssetAttributesTable()
    Dim lo As ListObject
    Set lo = Me.ListObjects("ClientAssetAttributes")
    If lo.DataBodyRange Is Nothing Then Exit Sub ' Already empty
    lo.DataBodyRange.Rows.Delete
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub

