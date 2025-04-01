VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub populateSF()

    Dim rPos(40) As Range, rThermo(40) As Range, rLow(40) As Range, rHigh(40) As Range, rAvg(40) As Range
    Dim i As Integer
    Dim sC As String, sP As String, sT As String, sL As String, sH As String, sA As String, sZ As String
    Dim sRLow As String, sRHigh As String, sRA
    Dim wb As Workbook
    Dim ws As Worksheet
    
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Initialize the individual cell variables for output grid
    i = 1
    Do While i <= 40
        
        Set rPos(i) = Cells(1, 25 + (i * 5))
        Set rThermo(i) = Cells(1, 26 + (i * 5))
        Set rLow(i) = Cells(1, 27 + (i * 5))
        Set rHigh(i) = Cells(1, 28 + (i * 5))
        Set rAvg(i) = Cells(1, 29 + (i * 5))
        
        If i < 10 Then
            sZ = "0"
        Else
            sZ = ""
        End If
                
        rPos(i).Value = "NV0" & sZ & i & "__c"
        rThermo(i).Value = "ResultText0" & sZ & i & "__c"
        rLow(i).Value = "Result0" & sZ & i & "__c"
        rHigh(i).Value = "Tol0" & sZ & i & "__c"
        rAvg(i).Value = "Avg0" & sZ & i & "__c"
        
        rPos(i).Offset(2, 0).Formula = "=IF(ISBLANK(Main!O" & 4 + i & "),"""",Main!N" & 4 + i & ")"
        rThermo(i).Offset(2, 0).Formula = "=IF(ISBLANK(Main!O" & 4 + i & "),"""",Main!O" & 4 + i & ")"
        rLow(i).Offset(2, 0).Formula = "=IF(Main!N" & 71 + i & "=0,"""", Round(Main!N" & 71 + i & ",1))"
        rHigh(i).Offset(2, 0).Formula = "=IF(Main!O" & 71 + i & "=0,"""", Round(Main!O" & 71 + i & ",1))"
        rAvg(i).Offset(2, 0).Formula = "=IF(Main!P" & 71 + i & "=0,"""", Round(Main!P" & 71 + i & ",1))"
        
        i = i + 1
    Loop
    

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
