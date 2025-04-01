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

Private Sub Worksheet_Calculate()

    'Check for errors ..
    On Error GoTo error_Alert
    
    'Define all variables
    Dim Col_set1(14) As Range, Col_set2(14) As Range, Col_set3(12) As Range, hy(40) As Range, ly(40) As Range, oPt(40) As Range, oPtClear(40) As Range
    Dim timePt(40) As Range, timeStp(40) As Range
    Dim updatedCell As Range, high As Range, tusTol As Range, low1st As Range, low2nd As Range, low3rd As Range, numPoints As Range
    Dim setPoint As Range, totalDevCell As Range, x1 As Range, x2 As Range, tempHold As Range
    Dim c As Integer, a_counter As Integer, h As Integer, ptColumn As Integer, colToSearch As Integer, iRow As Integer
    Dim lowOOT As Double, highOOT As Double, hold1 As Double, low As Double
    Dim highSign As Integer, lowSign As Integer
    Dim holdCell As Range, holdTime As Range
    Dim i As Double
    
    'Assign values to ranges
    Set low1st = Range("A34")
    Set low2nd = Range("U34")
    Set low3rd = Range("AO34")
    Set x1 = Range("D33")
    Set x2 = Range("H33")
    Set totalDevCell = Range("A52")
    Set numPoints = Range("B1")
    Set holdTime = Range("J1")

    'Check for empty field
    If IsError(x1) Then
        Exit Sub
    End If
    
    'Assign values to variables
    hold1 = Application.Sum(Range("A1:B2"), Range("B10:O25"), Range("V10:AI25"), Range("AP10:BA25"))
    highSign = Sgn(x1)
    lowSign = Sgn(x2)
    
    'Check to see if anything has changed **** Don't know if this is needed ****
    If hold1 = checkVal1 Then
        Exit Sub
    Else

        'Pre event settings
        Application.EnableEvents = False
        Application.ScreenUpdating = False
           
        'Assign values to ranges
        Set setPoint = Range("A1")
        Set high = Range("F34")
        Set tusTol = Range("A2")
        
        'Initialize high and low values
        lowOOT = setPoint - tusTol
        highOOT = setPoint + tusTol
        a_counter = 1
        colToSearch = 0
        
        'Get lowest temperature for the survey depending on the number of points
        If numPoints < 15 Then
            low = WorksheetFunction.Min(low1st)
        End If
        
        If numPoints > 14 And numPoints < 29 Then
            low = WorksheetFunction.Min(low1st, low2nd)
        End If
        
        If numPoints > 28 Then
            low = WorksheetFunction.Min(low1st, low2nd, low3rd)
        End If
       
        'Set holding variables to current highs and lows
        checkVal1 = hold1
        
        'Initialize the individual cell variables for output grid
        c = 1
        Do While c <= 40
            Set hy(c) = Range("J" & 36 + (c))
            Set ly(c) = Range("L" & 36 + (c))
            Set oPt(c) = Range("K" & 36 + (c))
            Set timePt(c) = Range("M" & 36 + (c))
            
            c = c + 1
        Loop
        
        'Initialize the TimeStamp locations and populate the variables
        c = 0
        Do While c < 17
            Set timeStp(c) = Range("A" & 10 + (c))
            c = c + 1
        Loop
        
        'Clear values over number of points
        Do While a_counter <= 40
            h = a_counter
            oPt(h).Value = ""
            timePt(h).Value = ""
            a_counter = a_counter + 1
        Loop
        
        'Process individual values
        a_counter = 1
        Do While a_counter <= numPoints
            h = a_counter
            ptColumn = 0
            
            'Reset Flag to Zero
            oPt(h).Value = 0
            timePt(h).Value = ""
            
            'Mark if High or Low in Tolerance
            If (hy(h) = "" Or ly(h) = "") Then
                oPt(h).Value = "Dropped"
            Else
                If Round(hy(h), 1) = Round(high, 1) Then
                    oPt(h).Value = "High"
                End If
                
                If Round(ly(h), 1) = Round(low, 1) Then
                    oPt(h).Value = "Low"
                End If
                
                'Mark if High or Low out of Tolerance
                If Round(ly(h), 1) < lowOOT Then
                    oPt(h).Value = "Out of Tol - LOW"
                End If
                
                If Round(hy(h), 1) > highOOT Then
                    oPt(h).Value = "Out of Tol - HIGH"
                End If
            
                'Get the time from the High or Low Temp for the given point
                If Right(oPt(h), 3) = "Low" Or Right(oPt(h), 4) = "High" Then
                    'Get Column Information
                    colToSearch = oPt(h).Row - 36
    
                    If colToSearch < 15 Then
                        ptColumn = h + 1
                    End If
                    
                    If colToSearch > 14 And h < 29 Then
                        ptColumn = h + 7
                    End If
                    
                    If colToSearch > 28 Then
                        ptColumn = h + 13
                    End If
                    
                    'Get Low Time
                    If Right(oPt(h), 3) = "Low" Then
                        For iRow = 10 To 25
                            If Not IsEmpty(Cells(iRow, ptColumn)) Then
                                Set tempHold = Cells(iRow, ptColumn)
                                If Round(tempHold, 1) = Round(low, 1) Then
                                    timePt(h) = timeStp(tempHold.Row - 10)
                                End If
                            End If
                        Next iRow
                    End If
                    
                    'Get High Time
                    If Right(oPt(h), 4) = "High" Then
                        For iRow = 10 To 25
                            If Not IsEmpty(Cells(iRow, ptColumn)) Then
                                Set tempHold = Cells(iRow, ptColumn)
                                If Round(tempHold, 1) = Round(high, 1) Then
                                    timePt(h) = timeStp(tempHold.Row - 10)
                                End If
                            End If
                        Next iRow
                    End If
                Else
                    timePt(h).Value = ""
                End If
            End If
            
            a_counter = a_counter + 1
        Loop
        
        'Process total deviation
        totalDevCell = 0
        If highSign = -1 And lowSign = -1 Then
            totalDevCell = x1 - x2
        Else
            If highSign = 1 And lowSign = 1 Then
                totalDevCell = x1 - x2
            Else
                totalDevCell = Abs(x1) + Abs(x2)
            End If
        End If
    End If
        
    'Reset all normal functions
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
error_Alert:
    If Err Then
        'MsgBox Err.Description
        'MsgBox Err.Description & Chr(13) & Err.Number
        
        'Reset Worksheet
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
End Sub



