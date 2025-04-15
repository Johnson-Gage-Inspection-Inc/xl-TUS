VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Write_Wire_Correction_Factors()

    'Check for errors
    On Error GoTo HandleError
    
    'Stop Screen Updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'Define all variables
    Dim MyDictionary As Object
    Dim intTestTemp As Long, intLowTemp As Long, intHighTemp As Long, c As Long, intLowCellNum As Long, intLastPoint As Long
    Dim intHighCellNum As Long, d As Long, iColumn As Long, iRow As Long, iHoldR As Long, iHoldC As Long
    Dim i As Long, iNumOfTestPoints As Long, iNumPointsTotalInWireData As Long
    Dim rCF() As Range, rTempTested() As Range, rWireLot() As Range, rWireLotMain(6) As Range, rNumPointsMain(6) As Range
    Dim rWirePoints(40) As Range
    Dim rLowTemp As Range, rHighTemp As Range, rCheckTemps As Range
    Dim wsStandards As Worksheet, wsMain As Worksheet, wsTUSWork As Worksheet
    Dim dLowCF As Double, dHighCF As Double, dTempSpread As Double, dCFSpread As Double, dFactor As Double, dCF As Double
        
    'Assign Variables
    Set wsStandards = ThisWorkbook.Worksheets("Standards_Import")
    Set wsMain = ThisWorkbook.Worksheets("Main")
    Set wsTUSWork = ThisWorkbook.Worksheets("TUS_Worksheet")
    Set MyDictionary = CreateObject("Scripting.Dictionary")
    intTestTemp = Me.Range("J7")
    iNumOfTestPoints = wsMain.Range("D17")
    
    'Clear old data
    Me.Range("A11:E27").ClearContents
    
    'Get number of columns
    iColumn = 0
    For i = 0 To 9
        If wsStandards.Cells(1, 3 + i) <> "" Then
            iColumn = iColumn + 1
        End If
    Next
    
    'Get number of rows
    iRow = wsStandards.Cells(Columns.count, 1).End(xlUp).Row - 2
    
    'Re Dim variables to proper sizes
    ReDim rWireLot(iRow)
    ReDim rCF(iRow)
    ReDim rTempTested(iColumn)
        
    'Assign Ranges
    For c = 0 To iRow - 1
        Set rWireLot(c) = wsStandards.Range("A" & 3 + c)
        Set rCF(c) = Me.Range("D" & 11 + c)
    Next c
    
    For c = 0 To iColumn - 1
        Set rTempTested(c) = wsStandards.Cells(1, 3 + c)
    Next c
    
    iNumPointsTotalInWireData = 0
    For c = 0 To 5
        Set rWireLotMain(c) = wsMain.Cells(55, 4 + c)
        Set rNumPointsMain(c) = wsMain.Cells(56, 4 + c)
        iNumPointsTotalInWireData = iNumPointsTotalInWireData + rNumPointsMain(c)
    Next
    
    ''**************Error Checking - Number of Points*********************
    If iNumPointsTotalInWireData <> iNumOfTestPoints Then
        MsgBox "The number of points used does not equal the number of test points in the survey."
        wsMain.Range("J56").Value = "Wire Usage Does NOT Equal Number of Test Points - Message will go away when you Re-Run the Data Import"
        Exit Sub
    Else
        wsMain.Range("J56").ClearContents
    End If
    
    'Clear old information
    wsTUSWork.Range("C12:P12").ClearContents
    wsTUSWork.Range("C18:P18").ClearContents
    wsTUSWork.Range("C24:P24").ClearContents
    
    For c = 0 To iNumOfTestPoints - 1
        If c < 14 Then
            Set rWirePoints(c) = wsTUSWork.Cells(12, 3 + c)
        End If

        If c > 13 And c < 28 Then
            Set rWirePoints(c) = wsTUSWork.Cells(18, 3 + (c - 14))
        End If

        If c > 27 Then
            Set rWirePoints(c) = wsTUSWork.Cells(24, 3 + (c - 28))
        End If
    Next
    
    'Get cell location of lower and upper range
    If intTestTemp < rTempTested(0).Value Then
        MsgBox rTempTested(0).Value & " is below the lowest certified temp for the wires you have chosen. Choose different wires."
        Resume CleanExit
    End If

    For i = 0 To iColumn - 1
        If intTestTemp >= rTempTested(i).Value Then
            intLowCellNum = rTempTested(i).Column
            intLowTemp = rTempTested(i).Value
            intHighCellNum = rTempTested(i).Offset(0, 1).Column
            intHighTemp = rTempTested(i).Offset(0, 1).Value
        End If
    Next i

    'Calculate Correction Factor
    For i = LBound(rWireLot) To (UBound(rWireLot) - 1)
        'Label the Wirelot in New Interp Sheet
        Me.Range("B" & 11 + i).Value = rWireLot(i)
        
        'Get the Values for the wirelot (Low Temp CF and High Temp CF)
        dLowCF = wsStandards.Cells(rWireLot(i).Row, intLowCellNum)
        dHighCF = wsStandards.Cells(rWireLot(i).Row, intHighCellNum)
        
        'Do the Math
        dTempSpread = intHighTemp - intLowTemp
        dCFSpread = dHighCF - dLowCF
        dFactor = dCFSpread / dTempSpread
        dCF = WorksheetFunction.Round(dLowCF + (intTestTemp - intLowTemp) * dFactor, 8)
        
        'Write the Correction Factor
        rCF(i) = WorksheetFunction.Round(dCF, 1)
        
        'Write to the dictionary (map)
        MyDictionary.Add Trim(rWireLot(i)), rCF(i)
    Next i
    
    'Populate Tus Worksheet with Correction Factors
    intLastPoint = 0
    For c = 0 To 5
        If rWireLotMain(c) <> "" Then
            For d = 0 To rNumPointsMain(c) - 1
                rWirePoints(d + intLastPoint) = MyDictionary.Item(Trim(rWireLotMain(c)))
            Next
            intLastPoint = intLastPoint + d
        End If
    Next
    
    'Call Daqbook Correction Sub
    Call Write_Daqbook_Correction_Factors

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox "Wire Correction " & Err.Description
    Resume CleanExit
End Sub

Private Sub Write_Daqbook_Correction_Factors()

    'Define all variables
    Dim intTestTemp As Long, intLowTemp As Long, intHighTemp As Long, c As Long, intLowCellNum As Long, intLastPoint As Long, intMidCellNum As Long, intMidTemp As Long
    Dim intHighCellNum As Long, d As Long, i As Long, iNumOfTestPoints As Long, cfTemp As Long, cfTempCell As Long
    Dim rCF() As Range, rTempTested(6) As Range, rTempMidTested(5) As Range
    Dim rDBPoints(40) As Range
    Dim wsStandards As Worksheet, wsMain As Worksheet, wsTUSWork As Worksheet
    Dim dTempSpreadMid As Double, dCF As Double

    'Assign Variables
    Set wsStandards = ThisWorkbook.Worksheets("Standards_Import")
    Set wsMain = ThisWorkbook.Worksheets("Main")
    Set wsTUSWork = ThisWorkbook.Worksheets("TUS_Worksheet")
    intTestTemp = wsMain.Range("D16")
    iNumOfTestPoints = wsMain.Range("D17")

    
    For c = 0 To 5
        'Put test temps in array
        Set rTempTested(c) = wsStandards.Cells(1, 16 + c)
        'Put midpoints into array
        Set rTempMidTested(c) = wsStandards.Cells(1, 26 + c)
        'Find temp range as compared to test temps
        If intTestTemp > rTempTested(c) Then
            intLowCellNum = rTempTested(c).Column
            intLowTemp = rTempTested(c).Value
            intHighCellNum = rTempTested(c).Offset(0, 1).Column
            intHighTemp = rTempTested(c).Offset(0, 1).Value
            intMidCellNum = rTempMidTested(c + 1).Column
            intMidTemp = rTempMidTested(c + 1).Value
        End If
    Next
    
    'Find Midpoint and then calculate closest temp
    dTempSpreadMid = (intHighTemp - intLowTemp) / 2
    
    If intTestTemp = (dTempSpreadMid + intLowTemp) Then
        cfTemp = intMidTemp
        cfTempCell = intMidCellNum
    Else
        If intTestTemp > (dTempSpreadMid + intLowTemp) Then
            cfTemp = intHighTemp
            cfTempCell = intHighCellNum
        Else
            cfTemp = intLowTemp
            cfTempCell = intLowCellNum
        End If
    End If
        
    'Clear old information
    wsTUSWork.Range("C11:P11").ClearContents
    wsTUSWork.Range("C17:P17").ClearContents
    wsTUSWork.Range("C23:P23").ClearContents
    
    'Write the CF to the TUS Survey worksheet
    For c = 0 To 39
        If c < 14 Then
            Set rDBPoints(c) = wsTUSWork.Cells(11, 3 + c)
        End If

        If c > 13 And c < 28 Then
            Set rDBPoints(c) = wsTUSWork.Cells(17, 3 + (c - 14))
        End If

        If c > 27 Then
            Set rDBPoints(c) = wsTUSWork.Cells(23, 3 + (c - 28))
        End If
    Next
    
    'Write correction factor
    For i = 0 To iNumOfTestPoints - 1

        If wsMain.Shapes("Check Box 5").ControlFormat.Value = 1 Then
             'Put Correction factor
            dCF = wsStandards.Cells(16 + i, cfTempCell)
            rDBPoints(i).Value = WorksheetFunction.Round(dCF, 10)
        End If
        
        If wsMain.Shapes("Check Box 6").ControlFormat.Value = 1 Then
             'Put Correction factor
            dCF = wsStandards.Cells(30 + i, cfTempCell)
            rDBPoints(i).Value = WorksheetFunction.Round(dCF, 10)
        End If
        
        If wsMain.Shapes("Check Box 6").ControlFormat.Value <> 1 And wsMain.Shapes("Check Box 5").ControlFormat.Value <> 1 Then
            'Put Correction factor
            dCF = wsStandards.Cells(2 + i, cfTempCell)
            rDBPoints(i).Value = WorksheetFunction.Round(dCF, 10)
        End If
        
        If wsMain.Shapes("Check Box 6").ControlFormat.Value = 1 And wsMain.Shapes("Check Box 5").ControlFormat.Value = 1 Then
            MsgBox "You can not check both CF offset boxes at the same time. Please uncheck one or both boxes and click Recalculate Correction Factors Button."
            Exit For
        End If
        
    Next

End Sub
