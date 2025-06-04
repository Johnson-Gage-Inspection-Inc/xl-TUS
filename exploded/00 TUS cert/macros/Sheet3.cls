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
    If Worksheets("Data_Sheet").Range("B16") = "" Then Exit Sub
    
    'Pre event settings
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Check for errors
    On Error GoTo HandleError
    
    'Clear the underlining
    Worksheets("Data_Sheet").Range("B12:P903").Font.Underline = False
    Worksheets("Data_Sheet_15_28").Range("B12:P903").Font.Underline = False
    Worksheets("Data_Sheet_29_40").Range("B12:P903").Font.Underline = False
    
    'Assign the starting row and ending row
    Dim x1 As Long, x2 As Long
    x1 = CLng(Mid(Range("P1").Value, 2))
    x2 = CLng(Mid(Range("P2").Value, 2))

    'Underline the survey data
    Dim i As Long
    For i = x1 To x2
        Worksheets("Data_Sheet").Rows(i).Font.Underline = True
        Worksheets("Data_Sheet_15_28").Rows(i).Font.Underline = True
        Worksheets("Data_Sheet_29_40").Rows(i).Font.Underline = True
    Next i

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub
















