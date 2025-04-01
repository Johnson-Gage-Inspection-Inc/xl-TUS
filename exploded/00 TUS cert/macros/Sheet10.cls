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
    Dim c As Collection
    Dim i As Integer
    Dim iLastRowCompany As Integer, iLastRowInsp As Integer, iLastRowField As Integer, iCompanyListPointer As Integer
    Dim v As Variant
    Dim r As Range, rSourceCompanyNames As Range, rSourceInspNames As Range
    Dim rTempList As Range, rCompanyList As Range, rHandheldList As Range
    Dim rTargetCompanyCell As Range, rTargetInspCell As Range, rTargetCompanyDD As Range, rTargetInspDD As Range
    Dim sSearchString As String, sRangeCompany As String, sRangeInsp As String
    Dim sRangeFlukes As String, sRangeDaqbooks As String, sRangeWires As String
    Dim wb As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    
    'Set Workbook/Worksheet variables
    Set wb = ThisWorkbook
    Set wsSource = wb.Worksheets("Furnace_Data")
    Set wsTarget = wb.Worksheets("Main")
    
    'Set Ranges
    Set rTargetCompanyCell = wsSource.Range("B2")
    Set rTargetInspCell = wsSource.Range("N2")
    Set rTargetCompanyDD = wsTarget.Range("D5")
    Set rTargetInspDD = wsTarget.Range("D7")
        
    'Initization Settings
    On Error GoTo HandleError
    
    'Check to see if the list has changed
    If Not Intersect(Target, Columns(34)) Is Nothing Then
    
        'Stop Events
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        'Populate Target Company Cell
        If rTargetCompanyCell.Value <> "" Or Not (IsEmpty(rTargetCompanyCell.Value)) Then
            
            'Get number of rows
            iLastRowCompany = wsSource.Range("B2", wsSource.Range("B2").End(xlDown)).Rows.count
            
            'Set ranges
            Set rSourceCompanyNames = wsSource.Range("B2:B" & iLastRowCompany + 1)
            Set c = New Collection
            
            'Continue on if there is a duplicate key
            On Error Resume Next
            
            'Set pointer
            iCompanyListPointer = 0
            
            'Process through the list of companies to create the list for the drop down
            For Each r In rSourceCompanyNames
                
                'Get the value of the cell
                v = r.Value
                
                'Check to see if v is empty
                If v <> "" Then
                    
                    'Add value and create key using string
                    c.Add v, CStr(v)
                    
                    'As long as error is 0 continue
                    If Err.Number = 0 Then
                    
                        'Increment Integer
                        iCompanyListPointer = iCompanyListPointer + 1
                        
                        'Populate Company List for Drop Down
                        wsSource.Range("K" & iCompanyListPointer).Value = CStr(v)
                        wsSource.Range("J" & iCompanyListPointer).Value = r.Offset(, 1).Value
                                        
                    Else
                    
                        'Reset error number to 0
                        Err.Number = 0
                    End If
                End If
            Next r
            
            'Start the error handling again
            On Error GoTo 0
            
            'Create Company Named Range
            sRangeCompany = "CompanyList"
            Set rCompanyList = wsSource.Range("K1:K" & iCompanyListPointer)
            wb.Names.Add Name:=sRangeCompany, RefersTo:=rCompanyList
            
            wsTarget.Range("D5:H5").ClearContents: wsTarget.Range("D6:H6").ClearContents: rTargetCompanyDD.Validation.Delete
            
            With rTargetCompanyDD.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="=CompanyList"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = False
            End With
        End If 'Populate Target Company Cell
        
        'Populate Inspector Cell
        If rTargetInspCell <> "" Or Not IsEmpty(rTargetInspCell) Then
            
            'Get number of rows
            iLastRowInsp = wsSource.Range("N2", wsSource.Range("N2").End(xlDown)).Rows.count
            
            'Set ranges
            Set rSourceInspNames = wsSource.Range("O2:O" & iLastRowInsp + 1)
            
            'Create Inspector Named Range
            sRangeInsp = "InspectorList"
            wb.Names.Add Name:=sRangeInsp, RefersTo:=rSourceInspNames
    
            If rTargetInspDD = "" Then
                wsTarget.Range("D7").ClearContents: rTargetInspDD.Validation.Delete
            
                With rTargetInspDD.Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="=InspectorList"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = False
                End With
            End If
        End If 'Populate Inspector Cell

        
    End If
    

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

