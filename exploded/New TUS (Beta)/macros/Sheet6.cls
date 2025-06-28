VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    'Setup Worksheet
    ' On Error GoTo HandleError

    'Initialize variables
    Dim r, aCell As Range, wireSetForSF(6) As Range, searchedSFId(6) As Range
    Dim sTmp As String, sArr() As String
    Dim v As Variant
    Dim i As Long, iHold As Long

    Dim wbTarget As Workbook:   Set wbTarget = ThisWorkbook
    Dim c As Collection:        Set c = New Collection
    Dim sDDString As String:    sDDString = ""
    Dim wsTarget As Worksheet:  Set wsTarget = wbTarget.Worksheets("Main")
    Dim rFurnaceDD As Range:    Set rFurnaceDD = wsTarget.Range("D6")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Daqbook Tag
    If Not Intersect(Target, Range("D9")) Is Nothing Then
        Call Sheet38.UpdateDaqbookServiceRecords
        ThisWorkbook.Connections("Query - DaqbookOffsets").Refresh
    End If
    
    ' Work Item Number
    If Not Intersect(Target, Range("K14")) Is Nothing Then
        Call Sheet25.UpdateWorkItemDetails
    End If
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

