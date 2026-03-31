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

Dim lastRadioVal As Variant

Private Sub Worksheet_Calculate()
    ' Intentionally empty - K14 trigger moved to Worksheet_Change
    ' to avoid calc-event suppression after synchronous PQ refresh.
End Sub

Private Sub Worksheet_Activate()
    ' Initialize memory on activation with non-error scalars
    Dim v As Variant

    v = Me.Range("L20").Value2
    If IsError(v) Then v = vbNullString
    lastRadioVal = v
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    ' K14 (JGI Order #) changed ? refresh WorkItemsCache
    If Not Intersect(Target, Me.Range("K14")) Is Nothing Then
        On Error GoTo HandleError
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Sheet25.UpdateWorkItemDetails
        GoTo CleanExit
    End If

    'Setup Worksheet
    On Error GoTo HandleError

    'Initialize variables
    Dim r, aCell As Range, wireSetForSF(6) As Range, searchedSFId(6) As Range
    Dim sTmp As String, sArr() As String
    Dim v As Variant
    Dim i As Long, iHold As Long

    Dim wbTarget As Workbook:   Set wbTarget = ThisWorkbook
    Dim c As Collection:        Set c = New Collection
    Dim sDDString As String:    sDDString = ""
    Dim rFurnaceDD As Range:    Set rFurnaceDD = Me.Range("D6")
    
    Application.ScreenUpdating = False

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True  ' Just in case
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub
