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
    On Error GoTo HandleError
    
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

    ' Company Name
    ' If Not Intersect(Target, Range("D5")) Is Nothing Then
    '     Call Sheet15.UpdateClientAssets(Range("D5").Value)
    ' End If
    
    ' Daqbook Tag
    If Not Intersect(Target, Range("D9")) Is Nothing Then
        Call Sheet38.UpdateDaqbookServiceRecords
    End If
    
    ' Work Item Number
    If Not Intersect(Target, Range("K14")) Is Nothing Then
        Call Sheet25.UpdateServiceOrderItems
    End If
    
    ' Client Asset Id
    If Not Intersect(Target, Range("D6:H6")) Is Nothing Then
        Call Sheet22.UpdateClientAssetAttributes
    End If

    If Not Intersect(Target, Range("D48:I48")) Is Nothing Then

        'Set variables and populate variables
        Range("L47:L52").ClearContents

        For i = 0 To 5
            Set wireSetForSF(i) = Range("L" & 47 + (i))
            Set searchedSFId(i) = Range("K" & 47 + (i))
            If searchedSFId(i) <> "" And InStr(sTmp, Trim(searchedSFId(i))) = 0 Then
                sTmp = sTmp & searchedSFId(i) & "-"
            End If
        Next i

        'Create Array with SF IDs
        If Len(sTmp) > 0 Then sTmp = Left(sTmp, Len(sTmp) - 1)
        sArr = Split(sTmp, "-")

        'Put IDs in slots
        For i = 0 To UBound(sArr)
            wireSetForSF(i).Value = sArr(i)
        Next i

    End If
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

