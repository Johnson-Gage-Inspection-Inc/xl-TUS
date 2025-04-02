Attribute VB_Name = "Module1"
Option Explicit

Sub OkEvents()
    Application.EnableEvents = True
End Sub

'Get Info for Drive
Function GetRootDrive(Optional aPath As String) As String
    GetRootDrive = CreateObject("Scripting.FileSystemObject").GetDriveName(aPath)
End Function

'User Defined Function for checking if a variable is in an array
Function IsInArray(valueToFind As Variant, arr As Variant) As Boolean
  On Error Resume Next
  IsInArray = Not IsError(Application.Match(valueToFind, arr, 0))
End Function

Function QRound(num_to_round As Double) As Double
    ' Banker's rounding (Round to even)
    QRound = Round(num_to_round, 1)
End Function

Sub Read_External_Workbook()

    'Check for errors
    On Error GoTo HandleError
    
    'Stop Screen Updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
     
    'Define Object for Target Workbook
    Dim strPath As String, strFile(6) As String, strFileDeDupped(6) As String
    Dim strFileHold1 As String, strFileHold2 As String, strFileHold3 As String, strDBFileName As String
    Dim rDBDate As Range, rDBName As Range, rDBPointsS() As Range, rDBPointsT() As Range, rDBCertTestTemp(6) As Range, rDBPointNum(40) As Range
    Dim wireLotRange(6) As Range
    Dim wireLot(6) As String, wireLotLetter(6) As String, wireLotNames(6) As String, wireLotNumber(6) As String
    Dim charNumber As Long, charNumberB As Long, charNumberA As Long, d As Long, iDBNumPoints As Long, iDBCertTemps(6) As Long
    Dim c As Long, intWireLotAmt As Long, intDistWireLotAmt As Long, i As Long, x As Long, intDistNumFileNames As Long, intDistNumUsedWireLots As Long
    Dim k As Long
    Dim bWireLotMatch As Boolean
    Dim v As Variant, v2 As Variant
    Dim obj As Object, obj2 As Object
    Dim Source_Workbook As Workbook, Source_Worksheet As Worksheet, Target_Workbook As Workbook, Target_Worksheet As Worksheet
            
    'Assign the Workbook Path
    strPath = GetRootDrive(ThisWorkbook.path) & "\Wires_Daqbook\"
    
    'Set Target Workbook
    Set Target_Workbook = ThisWorkbook
    Set Target_Worksheet = Target_Workbook.Worksheets("Standards_Import")
    
    'Initialize the wirelot information from Main page
    Set wireLotRange(0) = Worksheets("Main").Cells(55, 4)
    Set wireLotRange(1) = Worksheets("Main").Cells(55, 5)
    Set wireLotRange(2) = Worksheets("Main").Cells(55, 6)
    Set wireLotRange(3) = Worksheets("Main").Cells(55, 7)
    Set wireLotRange(4) = Worksheets("Main").Cells(55, 8)
    Set wireLotRange(5) = Worksheets("Main").Cells(55, 9)
    
    '=======================================================================================
    'Start Daqbook Data ====================================================================

    'Initialize the DaqBook information from Main Page
    Set rDBName = Worksheets("Main").Range("D9")
    Set rDBDate = Worksheets("Main").Range("D14")
    
    Select Case rDBName
        Case "J1"
            iDBNumPoints = Worksheets("Standards_Info").Range("E2")
        Case "J2"
            iDBNumPoints = Worksheets("Standards_Info").Range("E3")
        Case "J3"
            iDBNumPoints = Worksheets("Standards_Info").Range("E4")
        Case "K2"
            iDBNumPoints = Worksheets("Standards_Info").Range("E5")
        Case "K3"
            iDBNumPoints = Worksheets("Standards_Info").Range("E6")
        Case "K4"
            iDBNumPoints = Worksheets("Standards_Info").Range("E7")
        Case "K5"
            iDBNumPoints = Worksheets("Standards_Info").Range("E8")
        Case "K6"
            iDBNumPoints = Worksheets("Standards_Info").Range("E9")
        Case "N1"
            iDBNumPoints = Worksheets("Standards_Info").Range("E10")
        Case "N2"
            iDBNumPoints = Worksheets("Standards_Info").Range("E11")
    End Select
    
    'Set Array Size for Cells on source and target worksheets
    ReDim rDBPointsS(iDBNumPoints)
    ReDim rDBPointsT(iDBNumPoints)
    
    'Get filename for Daqbook and check and make sure file exists
    strDBFileName = strPath & rDBDate & ".xlsm"
    
    '**************Error Checking - Can't find Daqbook file*********************
    If Len(Dir(strDBFileName)) <> 0 Then
        strFile(c) = Trim(strFileHold1)
    Else
        MsgBox "A file can not be found for DaqBook " & rDBName & " with a certification date of " & rDBDate & "." & vbCrLf & "Please check the sub-directory and make sure the file exists."
        Exit Sub
    End If
    
    'Clear all cells before we import information
    Target_Worksheet.Columns("O:U").ClearContents
    
    'Open Daqbook File
    Set Source_Workbook = Workbooks.Open(strDBFileName)
    Set Source_Worksheet = Source_Workbook.Worksheets("Sheet1")
    
    'Get Test Temps and process CF from Source
    For i = 0 To 5
        'Label the spreadsheet with the correct temp and also assign the temp to the correct processing value
        iDBCertTemps(i) = Source_Worksheet.Range("A" & 42 + i).Value
        Set rDBCertTestTemp(i) = Target_Worksheet.Cells(1, 16 + i)
        rDBCertTestTemp(i) = iDBCertTemps(i)
    
        'Process each channel per given test temp
        For c = 0 To (iDBNumPoints - 1)
        
            'Label the Points
            Set rDBPointNum(c) = Target_Worksheet.Cells(2 + c, 15)
            rDBPointNum(c) = c + 1
            
            'Get the values
            If c < 6 Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet.Cells(42 + i, 2 + x)
                    Set rDBPointsT(c) = Target_Worksheet.Cells(2 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
    
            If c > 5 And c < 12 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet.Cells(50 + i, 2 + x)
                    Set rDBPointsT(c) = Target_Worksheet.Cells(8 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 11 And c < 18 And c < iDBNumPoints Then
                'If number of points is 14
                If iDBNumPoints = 14 Then
                    For x = 0 To 1
                        Set rDBPointsS(c) = Source_Worksheet.Cells(60 + i, 2 + x)
                        Set rDBPointsT(c) = Target_Worksheet.Cells(14 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                Else
                    'If number of points in not 14
                    For x = 0 To 5
                        Set rDBPointsS(c) = Source_Worksheet.Cells(60 + i, 2 + x)
                        Set rDBPointsT(c) = Target_Worksheet.Cells(14 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                End If
            End If
    
            If c > 17 And c < 24 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet.Cells(68 + i, 2 + x)
                    Set rDBPointsT(c) = Target_Worksheet.Cells(20 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 23 And c < 30 And c < iDBNumPoints Then
                If iDBNumPoints = 28 Then
                    'If number of points is 28
                    For x = 0 To 3
                        Set rDBPointsS(c) = Source_Worksheet.Cells(78 + i, 2 + x)
                        Set rDBPointsT(c) = Target_Worksheet.Cells(26 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                Else
                    'If number of points is not 28
                    For x = 0 To 5
                        Set rDBPointsS(c) = Source_Worksheet.Cells(78 + i, 2 + x)
                        Set rDBPointsT(c) = Target_Worksheet.Cells(26 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                End If
            End If
            
            If c > 29 And c < 36 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet.Cells(86 + i, 2 + x)
                    Set rDBPointsT(c) = Target_Worksheet.Cells(32 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 35 And c < iDBNumPoints Then
                For x = 0 To 3
                    Set rDBPointsS(c) = Source_Worksheet.Cells(96 + i, 2 + x)
                    Set rDBPointsT(c) = Target_Worksheet.Cells(38 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
        Next
    Next

    'Close Target Workbook
    Source_Workbook.Close False
    
    'END DaqBOOK Data ======================================================================
    '=======================================================================================
                
                
    '=======================================================================================
    'Start WireLot Data ====================================================================

    'Set holding ints to 0
    c = 0
    intWireLotAmt = 0
    

    'Check and see how many wirelots need to be processed
    Do While c < 6
        If wireLotRange(c).Value <> "" Then
            intWireLotAmt = intWireLotAmt + 1
            wireLotNames(c) = Trim(wireLotRange(c).Value)
        End If
        c = c + 1
    Loop

    'Remove duplicate lot numbers and create array for getting files
    Set obj = CreateObject("Scripting.Dictionary")

    For i = LBound(wireLotNames) To UBound(wireLotNames)
        obj(wireLotNames(i)) = 1
    Next i

    x = 0
    For Each v In obj.keys()
       wireLot(x) = v
       x = x + 1
    Next v
    
    'Set number of Used Wirelots
    intDistNumUsedWireLots = x - 1

    'Get number of distinct wirelots and assign them to variable
    intDistWireLotAmt = x - 1

    'Populate strFile Array
    c = 0
    Do While c < intDistWireLotAmt
        'get basic info for filename
        wireLotLetter(c) = UCase(Right(wireLot(c), 1))
        wireLotNumber(c) = Left(wireLot(c), 6)

        'Create possible combinations of the filename based off of the info
        charNumber = Asc(wireLotLetter(c))
        charNumberB = charNumber - 1
        charNumberA = charNumber + 1

        strFileHold1 = strPath & wireLotNumber(c) & Chr(charNumber) & ".xls"
        strFileHold2 = strPath & wireLotNumber(c) & Chr(charNumberB) & "-" & Chr(charNumber) & ".xls"
        strFileHold3 = strPath & wireLotNumber(c) & Chr(charNumber) & "-" & Chr(charNumberA) & ".xls"

        'check and see if the filename exists and what the actual filename is
        If Len(Dir(strFileHold1)) <> 0 Then
            strFile(c) = strFileHold1
        End If

        If Len(Dir(strFileHold2)) <> 0 Then
            strFile(c) = strFileHold2
        End If

        If Len(Dir(strFileHold3)) <> 0 Then
            strFile(c) = strFileHold3
        End If

        '**************Error Checking - Can't find wirelot file*********************
        If strFile(c) = "" Then
            MsgBox "A file can not be found for Wirelot Number " & wireLotNumber(c) & wireLotLetter(c) & "." & vbCrLf & "Please check the sub-directory and make sure the file exists."
            Exit Sub
        End If
        c = c + 1
    Loop

    'Remove Duplicate Filenames from list
    x = 0
    Set obj2 = CreateObject("Scripting.Dictionary")

    For i = LBound(strFile) To UBound(strFile)
        obj2(strFile(i)) = 1
    Next i

    For Each v2 In obj2.keys()
       strFileDeDupped(x) = v2
       x = x + 1
    Next v2

    'Get number of distinct file names
    intDistNumFileNames = x - 1

    'Clear all cells before we import information
    Target_Worksheet.Columns("A").ClearContents
    Target_Worksheet.Columns("C:M").ClearContents

    'Open Thermocouple file(s) and grab information
    c = 0
    d = 0
    Do While c < intDistNumFileNames
        If Not IsNull(strFileDeDupped(c)) Then

            Set Source_Workbook = Workbooks.Open(strFileDeDupped(c))
            Set Source_Worksheet = Source_Workbook.Worksheets("TC Form")

            'Check to see if this is the first run
            If c = 0 Then
                Target_Worksheet.Range("C1:G1").Value = Source_Worksheet.Range("D650:H650").Value
                Target_Worksheet.Range("H1:L1").Value = Source_Worksheet.Range("D656:H656").Value
            Else
                If bWireLotMatch = False Then
                    d = c
                Else
                    d = c + 1
                End If
            End If

            'Update Target File
            If IsInArray(Source_Worksheet.Range("B651").Value, wireLot) Then
                Target_Worksheet.Range("A" & 3 + (d)).Value = Source_Worksheet.Range("B651").Value
                Target_Worksheet.Range("C" & 3 + (d) & ":G" & 3 + (d)).Value = Source_Worksheet.Range("K653:O653").Value
                Target_Worksheet.Range("H" & 3 + (d) & ":L" & 3 + (d)).Value = Source_Worksheet.Range("K660:O660").Value
            Else
                d = d - 1
            End If
            
            If IsInArray(Source_Worksheet.Range("B691").Value, wireLot) And Source_Worksheet.Range("B691").Value <> 0 Then
                Target_Worksheet.Range("A" & 4 + (d)).Value = Source_Worksheet.Range("B691").Value
                Target_Worksheet.Range("C" & 4 + (d) & ":G" & 4 + (d)).Value = Source_Worksheet.Range("K693:O693").Value
                Target_Worksheet.Range("H" & 4 + (d) & ":L" & 4 + (d)).Value = Source_Worksheet.Range("K700:O700").Value
            Else
                bWireLotMatch = False
            End If
            
            'Close Target Workbook
            Source_Workbook.Close False
        End If
        c = c + 1
    Loop

    'END WireLot Data ======================================================================
    '=======================================================================================
    
    'Call the Wire Correction sub
    Call Write_Wire_Correction_Factors

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

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
    Dim ws As Worksheet, wsStandards As Worksheet, wsMain As Worksheet, wsTUSWork As Worksheet
    Dim dLowCF As Double, dHighCF As Double, dTempSpread As Double, dCFSpread As Double, dFactor As Double, dCF As Double
        
    'Assign Variables
    Set ws = ThisWorkbook.Worksheets("Interp")
    Set wsStandards = ThisWorkbook.Worksheets("Standards_Import")
    Set wsMain = ThisWorkbook.Worksheets("Main")
    Set wsTUSWork = ThisWorkbook.Worksheets("TUS_Worksheet")
    Set MyDictionary = CreateObject("Scripting.Dictionary")
    intTestTemp = ws.Range("J7")
    iNumOfTestPoints = wsMain.Range("D17")
    
    'Clear old data
    ws.Range("A11:E27").ClearContents
    
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
    c = 0
    Do While c < (iRow)
        Set rWireLot(c) = wsStandards.Range("A" & 3 + (c))
        Set rCF(c) = ws.Range("D" & 11 + (c))
        c = c + 1
    Loop
    
    c = 0
    Do While c < (iColumn)
        Set rTempTested(c) = wsStandards.Cells(1, 3 + c)
        c = c + 1
    Loop
    
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
    c = 0
    If intTestTemp < rTempTested(0).Value Then
        MsgBox rTempTested(0).Value & " is below the lowest certified temp for the wires you have chosen. Choose different wires."
        Resume CleanExit
    End If
    
    Do While c < (iColumn)
        If intTestTemp >= rTempTested(c) Then
            intLowCellNum = rTempTested(c).Column
            intLowTemp = rTempTested(c).Value
            intHighCellNum = rTempTested(c).Offset(0, 1).Column
            intHighTemp = rTempTested(c).Offset(0, 1).Value
        End If
        c = c + 1
    Loop

    'Calculate Correction Factor
    For i = LBound(rWireLot) To (UBound(rWireLot) - 1)
        'Label the Wirelot in New Interp Sheet
        ws.Range("B" & 11 + i).Value = rWireLot(i)
        
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
    Dim ws As Worksheet, wsStandards As Worksheet, wsMain As Worksheet, wsTUSWork As Worksheet
    Dim dTempSpreadMid As Double, dCF As Double

    'Assign Variables
    Set ws = ThisWorkbook.Worksheets("Interp")
    Set wsStandards = ThisWorkbook.Worksheets("Standards_Import")
    Set wsMain = ThisWorkbook.Worksheets("Main")
    Set wsTUSWork = ThisWorkbook.Worksheets("TUS_Worksheet")
    intTestTemp = wsMain.Range("D16")
    iNumOfTestPoints = wsMain.Range("D17")

    'Put test temps in array
    For c = 0 To 5
        Set rTempTested(c) = wsStandards.Cells(1, 16 + c)
    Next
    
    'Put midpoints into array
    For c = 1 To 5
        Set rTempMidTested(c) = wsStandards.Cells(1, 26 + c)
    Next
    
    'Find temp range as compared to test temps
    For c = 0 To 5
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


