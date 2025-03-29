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
    On Error GoTo Whoa
    
    'Define variables
    Dim rSourceNames As Range, rFurnaceDD As Range, r As Range, aCell As Range, rFurnaceList As Range
    Dim wireSetForSF(6) As Range, searchedSFId(6) As Range
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim sDDString As String, sSearchString As String, sRangeName As String
    Dim sFurnaceList As String, sTmp As String, sArr() As String
    Dim c As Collection
    Dim v As Variant
    Dim i As Integer, iHold As Integer
    Dim iLastRow As Integer
    
    'Set variables
    Set wbTarget = ThisWorkbook
    Set wsSource = wbTarget.Worksheets("Furnace_Data")
    Set wsTarget = wbTarget.Worksheets("Main")
    iLastRow = wsSource.Range("B2", wsSource.Range("B2").End(xlDown)).Rows.Count
    Set rSourceNames = wsSource.Range("B2:B" & iLastRow)
    Set rFurnaceDD = wsTarget.Range("D6")
    Set c = New Collection
    sDDString = ""
    
    '~~> Capturing change in cell D5
    If Not Intersect(Target, Range("D5")) Is Nothing Then
    
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        wsTarget.Unprotect Password:="JGIPyro"
        
        'Set variables
        sSearchString = Range("D5").value
        
        'Check to make sure there is a company selected, if not, reset worksheet
        If Trim$(sSearchString) = "" Then
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Worksheets("Main").Protect Password:="JGIPyro"
            Range("D6:H6").ClearContents
            Exit Sub
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
        wsTarget.Unprotect Password:="JGIPyro"
        
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
        iHold = UBound(sArr) + 1
        
        If iHold > 0 Then
        
            'Process Array
            i = 0
            Do While i < iHold
                wireSetForSF(i).value = sArr(i)
                i = i + 1
            Loop
            
        End If
        
    
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("Main").Protect Password:="JGIPyro"
    
LetsContinue:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("Main").Protect Password:="JGIPyro"
    Exit Sub
    
Whoa:
    MsgBox Err.Description
    Resume LetsContinue
    
End Sub

