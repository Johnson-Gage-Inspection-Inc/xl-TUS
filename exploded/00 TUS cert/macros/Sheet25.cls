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
Public Sub UpdateServiceOrderItems()
    TruncateServiceOrderItemsTable
    ThisWorkbook.Connections("Query - ServiceOrderItems").Refresh
End Sub
Private Sub TruncateServiceOrderItemsTable()
    Dim lo As ListObject
    Set lo = Me.ListObjects("ServiceOrderItems")
    If lo.DataBodyRange Is Nothing Then Exit Sub ' Already empty
    lo.DataBodyRange.Rows.Delete
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    ThisWorkbook.Connections("Query - AssetId").Refresh
    ThisWorkbook.Connections("Query - ClientCompanyId").Refresh
    Sheet27.UpdateClientAsset
    Sheet22.UpdateClientAssetAttributes
End Sub
