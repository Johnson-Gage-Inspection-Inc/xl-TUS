VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Calculate()
    
    'Check for data
    If Worksheets("Data_Sheet").Range("B16") = "" Then
        Exit Sub
    End If
    
    'Pre event settings
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Check for errors
    On Error GoTo HandleError
    
    'Define all variables
    Dim beginRow As Range, endRow As Range, changeRow(16) As Range
    Dim hold1 As Double, hold2 As Double
    Dim x1 As Long, x2 As Long, amountChar1 As Long, amountChar2 As Long
            
    'Assign cell locations to variables
    Set beginRow = Range("P1")
    Set endRow = Range("P2")
        
    'See how many char are in the address, is always A so this checks for two or three digits
    amountChar1 = Len(beginRow)
    amountChar2 = Len(endRow)
    
    'Clear the underlining
    Worksheets("Data_Sheet").Range("B12:P903").Font.Underline = False
    Worksheets("Data_Sheet_15_28").Range("B12:P903").Font.Underline = False
    Worksheets("Data_Sheet_29_40").Range("B12:P903").Font.Underline = False
    
    'Assign the starting row and ending row
    x1 = Right(beginRow, amountChar1 - 1)
    x2 = Right(endRow, amountChar2 - 1)

    'Underline the survey data
    Do While x1 <= x2
        Worksheets("Data_Sheet").Rows(x1).Font.Underline = True
        Worksheets("Data_Sheet_15_28").Rows(x1).Font.Underline = True
        Worksheets("Data_Sheet_29_40").Rows(x1).Font.Underline = True
        x1 = x1 + 1
    Loop
    
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

