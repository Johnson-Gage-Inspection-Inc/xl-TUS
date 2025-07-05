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

Private Sub Worksheet_Change(ByVal Target As Range)
    ThisWorkbook.Connections("Query - CorrectionFactors").Refresh
End Sub

Public Sub UpdateDaqbookServiceRecords()
    ClearValueColumnInDaqbookServiceRecordsTable
    ThisWorkbook.Connections("Query - DaqbookServiceRecord").Refresh
End Sub
Private Sub ClearValueColumnInDaqbookServiceRecordsTable()
    Dim lo As ListObject
    Dim valueCol As ListColumn
    Set lo = Me.ListObjects("DaqbookServiceRecord")
    
    On Error Resume Next
    Set valueCol = lo.ListColumns("Value")
    On Error GoTo 0
    
    If Not valueCol Is Nothing Then
        If Not valueCol.DataBodyRange Is Nothing Then
            valueCol.DataBodyRange.ClearContents
        End If
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call EnforceSheetIsViewOnly(Me)
End Sub
