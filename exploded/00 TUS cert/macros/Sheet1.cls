VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public checkVal1 As Double
Function SafeSum(rng As Range) As Double
    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            SafeSum = SafeSum + cell.Value
        Else
            SafeSum = 0
            Exit Function
        End If
    Next cell
End Function

Private Sub Worksheet_Calculate()

    'Check for errors ..
    ' On Error GoTo HandleError
    
    'Define all variables
    Dim hy(40) As Range, ly(40) As Range, oPt(40) As Range
    Dim timePt(40) As Range, timeStp(40) As Range, tempHold As Range
    Dim i As Long, ptColumn As Long, iRow As Long
    Dim low As Double
    
    Dim holdCell As Range
    
    'Assign values to ranges
    Dim low1st As Range:        Set low1st = Range("A34")
    Dim low2nd As Range:        Set low2nd = Range("U34")
    Dim low3rd As Range:        Set low3rd = Range("AO34")
    Dim x1 As Range:            Set x1 = Range("D33")
    Dim x2 As Range:            Set x2 = Range("H33")
    Dim totalDevCell As Range:  Set totalDevCell = Range("A52")
    Dim numPoints As Range:     Set numPoints = Range("B1")
    Dim holdTime As Range:      Set holdTime = Range("J1")

    'Check for empty field
    If IsError(x1) Then
        Exit Sub
    End If

    'Assign values to variables
    Dim hold1 As Double:
    hold1 = SafeSum(Range("A1:B2")) + SafeSum(Range("B10:O25")) + SafeSum(Range("V10:AI25")) + SafeSum(Range("AP10:BA25"))
    
    'Check to see if anything has changed **** Don't know if this is needed ****
    If hold1 = checkVal1 Then
        Exit Sub
    End If

    'Pre event settings
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Assign values to ranges
    Dim setPoint As Range:  Set setPoint = Range("A1")
    Dim high As Range:      Set high = Range("F34")
    Dim tusTol As Range:    Set tusTol = Range("A2")

    'Initialize high and low values
    Dim lowOOT As Double:       lowOOT = setPoint - tusTol
    Dim highOOT As Double:      highOOT = setPoint + tusTol
    Dim a_counter As Long:      a_counter = 1
    Dim colToSearch As Long:    colToSearch = 0
    
    'Get lowest temperature for the survey depending on the number of points
    If numPoints < 15 Then
        low = WorksheetFunction.Min(low1st)
    ElseIf numPoints > 14 And numPoints < 29 Then
        low = WorksheetFunction.Min(low1st, low2nd)
    ElseIf numPoints > 28 Then
        low = WorksheetFunction.Min(low1st, low2nd, low3rd)
    End If
   
    'Set holding variables to current highs and lows
    checkVal1 = hold1
    
    'Initialize the individual cell variables for output grid
    For i = 1 To 40
        Set hy(i) = Range("J" & 36 + i)
        Set ly(i) = Range("L" & 36 + i)
        Set oPt(i) = Range("K" & 36 + i)
        Set timePt(i) = Range("M" & 36 + i)
    Next i
    
    'Initialize the TimeStamp locations and populate the variables
    For i = 0 To 16
        Set timeStp(i) = Range("A" & (10 + i))
    Next i
    
    'Clear values over number of points
    For i = 1 To 40
        oPt(i).Value = ""
        timePt(i).Value = ""
    Next i

    'Process individual values
    For i = 1 To numPoints
        ptColumn = 0
        
        'Reset Flag to Zero
        oPt(i).Value = 0
        timePt(i).Value = ""
        
        'Mark if High or Low in Tolerance
        If (hy(i) = "" Or ly(i) = "") Then
            oPt(i).Value = "Dropped"
        Else
            
            'Mark if High or Low out of Tolerance
            If Round(ly(i), 1) < lowOOT Then
                oPt(i).Value = "Out of Tol - LOW"
            ElseIf Round(hy(i), 1) > highOOT Then
                oPt(i).Value = "Out of Tol - HIGH"
            End If

            If Round(hy(i), 1) = Round(high, 1) Then
                oPt(i).Value = "High"
            ElseIf Round(ly(i), 1) = Round(low, 1) Then
                oPt(i).Value = "Low"
            Else
                timePt(i).Value = ""
                GoTo NextPoint
            End If

            'Get Column Information
            colToSearch = oPt(i).row - 36

            If colToSearch < 15 Then
                ptColumn = i + 1
            ElseIf colToSearch > 14 And i < 29 Then
                ptColumn = i + 7
            ElseIf colToSearch > 28 Then
                ptColumn = i + 13
            End If
            
            If Right(oPt(i), 3) = "Low" Then
                timePt(i) = FindTimePoint(low, "Low", ptColumn, timeStp)
            ElseIf Right(oPt(i), 4) = "High" Then
                timePt(i) = FindTimePoint(high, "High", ptColumn, timeStp)
            End If

        End If
NextPoint:
    Next i
    
    'Process total deviation
    totalDevCell = IIf(Sgn(x1) = Sgn(x2), x1 - x2, Abs(x1) + Abs(x2))
        
CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

Function FindTimePoint(targetValue As Variant, suffix As String, col As Long, timeStp() As Range)
    Dim iRow As Long
    For iRow = 10 To 25
        If Not IsEmpty(Cells(iRow, col)) Then
            If Round(Cells(iRow, col).Value, 1) = Round(targetValue, 1) Then
                FindTimePoint = timeStp(iRow - 10)
                Exit Function
            End If
        End If
    Next iRow
End Function


