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
    Worksheets("Data_Sheet").Unprotect Password:="JGIPyro"
    Worksheets("Data_Sheet_15_28").Unprotect Password:="JGIPyro"
    Worksheets("Data_Sheet_29_40").Unprotect Password:="JGIPyro"
    
    'Check for errors
    On Error GoTo error_Alert
    
    'Define all variables
    Dim beginRow As Range, endRow As Range, changeRow(16) As Range
    Dim hold1 As Double, hold2 As Double
    Dim x1 As Integer, x2 As Integer, amountChar1 As Integer, amountChar2 As Integer
            
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
    
    'Post Event Settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Worksheets("Data_Sheet").Protect Password:="JGIPyro"
    Worksheets("Data_Sheet_15_28").Protect Password:="JGIPyro"
    Worksheets("Data_Sheet_29_40").Protect Password:="JGIPyro"
    
error_Alert:
    If Err Then
        'Output Error Message
        MsgBox Err.Description & Chr(13) & Err.Number
        
        'Reset Worksheet
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Worksheets("Data_Sheet").Protect Password:="JGIPyro"
        Worksheets("Data_Sheet_15_28").Protect Password:="JGIPyro"
        Worksheets("Data_Sheet_29_40").Protect Password:="JGIPyro"
    End If
End Sub

