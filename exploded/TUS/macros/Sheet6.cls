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
Dim lastWorkOrderKey As Variant

Private Sub Worksheet_Calculate()

    ' Watch G2 for GetFormattedWorkItemNumber(K14)
    Dim currentKey As Variant
    currentKey = Me.Range("G2").value
    
    ' If G2 is an error, bail
    If IsError(currentKey) Then GoTo CleanAfterCalc
    
    ' If our last cached value was an error (e.g., from a prior calc), reset it
    If VarType(lastWorkOrderKey) = vbError Then lastWorkOrderKey = vbNullString
    
    ' Safe compare
    If CStr(currentKey) <> CStr(lastWorkOrderKey) Then
        lastWorkOrderKey = currentKey
        On Error GoTo CleanAfterCalc
        Application.EnableEvents = False
        Sheet25.UpdateWorkItemDetails
    End If

CleanAfterCalc:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Activate()
    ' Initialize memory on activation with non-error scalars
    Dim v As Variant

    v = Me.Range("L20").Value2
    If IsError(v) Then v = vbNullString
    lastRadioVal = v

    v = Me.Range("G2").Value2
    If IsError(v) Then v = vbNullString
    lastWorkOrderKey = v
End Sub

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
