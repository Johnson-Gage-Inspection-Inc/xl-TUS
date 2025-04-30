VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Sub UpdateDaqbookServiceRecords()
    TruncateDaqbookServiceRecordsTable
    ThisWorkbook.Connections("Query - DaqbookServiceRecords").Refresh
End Sub
Private Sub TruncateDaqbookServiceRecordsTable()
    Dim lo As ListObject
    Set lo = Me.ListObjects("DaqbookServiceRecords")
    If lo.DataBodyRange Is Nothing Then Exit Sub ' Already empty
    lo.DataBodyRange.Rows.Delete
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
