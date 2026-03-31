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
    Application.ScreenUpdating = False

    ' Force synchronous refresh so data is available before Me.Calculate.
    ' The connection's BackgroundQuery property stays True for manual/RefreshAll use,
    ' but we temporarily override it here to avoid a race condition.
    If ConnectionExists("Query - WorkItemsCache") Then
        SyncRefresh ThisWorkbook.Connections("Query - WorkItemsCache")
    ElseIf ConnectionExists("Query - WorkItemDetails") Then
        ' Backward-compatibility fallback for older workbook snapshots.
        SyncRefresh ThisWorkbook.Connections("Query - WorkItemDetails")
    End If

    Me.Calculate
    Application.ScreenUpdating = True
End Sub

Private Sub SyncRefresh(ByVal conn As WorkbookConnection)
    Dim wasBg As Boolean
    On Error Resume Next
    wasBg = conn.OLEDBConnection.BackgroundQuery
    conn.OLEDBConnection.BackgroundQuery = False
    On Error GoTo 0

    conn.Refresh

    On Error Resume Next
    conn.OLEDBConnection.BackgroundQuery = wasBg
    On Error GoTo 0
End Sub

Private Function ConnectionExists(ByVal connectionName As String) As Boolean
    On Error Resume Next
    Dim conn As WorkbookConnection
    Set conn = ThisWorkbook.Connections(connectionName)
    ConnectionExists = Not conn Is Nothing
    Set conn = Nothing
    On Error GoTo 0
End Function

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
