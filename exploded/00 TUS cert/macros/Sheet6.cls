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
    Dim r, aCell As Range, rFurnaceList As Range, wireSetForSF(6) As Range, searchedSFId(6) As Range
    Dim sSearchString As String, sRangeName As String, sFurnaceList As String, sTmp As String, sArr() As String
    Dim v As Variant
    Dim i As Long, iHold As Long

    Dim wbTarget As Workbook:  Set wbTarget = ThisWorkbook
    Dim c As Collection:       Set c = New Collection
    Dim sDDString As String:   sDDString = ""
    Dim wsSource As Worksheet: Set wsSource = wbTarget.Worksheets("Furnace_Data")
    Dim iLastRow As Long:      iLastRow = wsSource.Range("B2").End(xlDown).Row - 1
    Dim rSourceNames As Range: Set rSourceNames = wsSource.Range("B2:B" & iLastRow)
    Dim wsTarget As Worksheet: Set wsTarget = wbTarget.Worksheets("Main")
    Dim rFurnaceDD As Range:   Set rFurnaceDD = wsTarget.Range("D6")
    
    '~~> Capturing change in cell D5
    If Not Intersect(Target, Range("D5")) Is Nothing Then
    
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        'Set variables
        sSearchString = Range("D5").Value
        
        'Check to make sure there is a company selected, if not, reset worksheet
        If Trim$(sSearchString) = "" Then
            Range("D6:H6").ClearContents
            GoTo CleanExit
        End If
        
        'Find the Range for th list
        sFurnaceList = FindRange(rSourceNames, sSearchString, 2)

        'Reactivate the Target worksheet
        wsTarget.Activate

        'Create Named Range
        sRangeName = "FurnaceList"
        Set rFurnaceList = wsSource.Range(sFurnaceList)
        wbTarget.Names.Add Name:=sRangeName, RefersTo:=rFurnaceList
        
        'Clear previous Drop Down list
        Range("D6:H6").ClearContents: rFurnaceDD.Validation.Delete
 
        '~~> Create the DV List
        With rFurnaceDD.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=FurnaceList"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
        
    End If
    
    If Not Intersect(Target, Range("D48:I48")) Is Nothing Then
    
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        'Set variables and populate variables
        Range("L47:L52").ClearContents
        i = 0
        Do While i < 6
            Set wireSetForSF(i) = Range("L" & 47 + (i))
            Set searchedSFId(i) = Range("K" & 47 + (i))
            
            If searchedSFId(i) <> "" Then
                
                If InStr(sTmp, Trim(searchedSFId(i))) = 0 Then
                    sTmp = sTmp & searchedSFId(i) & "-"
                End If
                
            End If
            
            i = i + 1
            
        Loop
        
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

