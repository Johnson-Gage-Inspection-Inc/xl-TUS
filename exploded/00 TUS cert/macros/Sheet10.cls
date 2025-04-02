VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    'Define Variables
    Dim r As Range
    Dim rTempList As Range, rCompanyList As Range, rHandheldList As Range

    'Set Workbook/Worksheet variables
    Dim wb As Workbook:         Set wb = ThisWorkbook
    Dim wsTarget As Worksheet:  Set wsTarget = wb.Worksheets("Main")

    'Set Ranges
    Dim rTargetCompanyCell As Range: Set rTargetCompanyCell = Me.Range("B2")
    Dim rTargetInspCell As Range:    Set rTargetInspCell = Me.Range("N2")
    Dim rTargetCompanyDD As Range:   Set rTargetCompanyDD = wsTarget.Range("D5")
    Dim rTargetInspDD As Range:      Set rTargetInspDD = wsTarget.Range("D7")

    'Initization Settings
    On Error GoTo HandleError

    'Check to see if the list has changed
    If Intersect(Target, Columns(34)) Is Nothing Then
        GoTo CleanExit
    End If

    'Stop Events
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Populate Target Company Cell
    If rTargetCompanyCell.Value <> "" Or Not (IsEmpty(rTargetCompanyCell.Value)) Then
        Dim iLastRowCompany As Integer:     iLastRowCompany = Me.Range("B2").End(xlDown).Row - 1
        Dim rSourceCompanyNames As Range:   Set rSourceCompanyNames = Me.Range("B2:B" & iLastRowCompany + 1)
        Dim c As Collection:                Set c = New Collection
        Dim iCompanyListPointer As Integer: iCompanyListPointer = 0
        Dim dict As Object:                 Set dict = CreateObject("Scripting.Dictionary") ' Late binding
        Dim v As String
        
        For Each r In rSourceCompanyNames
            v = Trim(r.Value)
            If Len(v) > 0 And Not dict.Exists(v) Then
                dict.Add v, r.Offset(0, 1).Value
                iCompanyListPointer = iCompanyListPointer + 1
                Me.Cells(iCompanyListPointer, "K").Value = v
                Me.Cells(iCompanyListPointer, "J").Value = r.Offset(0, 1).Value
            End If
        Next r

        'Create Company Named Range
        Set rCompanyList = Me.Range("K1:K" & iCompanyListPointer)
        wb.Names.Add Name:="CompanyList", RefersTo:=rCompanyList

        wsTarget.Range("D5:H6").ClearContents
        rTargetCompanyDD.Validation.Delete
        ApplyValidation rTargetCompanyDD, "=CompanyList"
    End If 'Populate Target Company Cell

    'Populate Inspector Cell
    If rTargetInspCell <> "" Or Not IsEmpty(rTargetInspCell) Then
        Dim iLastRowInsp As Integer:    iLastRowInsp = Me.Range("N2").End(xlDown).Row - 1
        Dim rSourceInspNames As Range:  Set rSourceInspNames = Me.Range("O2:O" & iLastRowInsp + 1)

        'Create Inspector Named Range
        wb.Names.Add Name:="InspectorList", RefersTo:=rSourceInspNames

        If rTargetInspDD = "" Then
            wsTarget.Range("D7").ClearContents: rTargetInspDD.Validation.Delete
            ApplyValidation rTargetInspDD, "=InspectorList"
        End If
    End If 'Populate Inspector Cell

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

Private Sub ApplyValidation(rng As Range, listName As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
End Sub

