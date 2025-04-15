Attribute VB_Name = "modMainLogic"
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
    Dim strFile(6) As String, strFileDeDupped(6) As String
    Dim strFileHold1 As String, strFileHold2 As String, strFileHold3 As String, strDBFileName As String
    Dim rDBDate As Range, rDBName As Range, rDBPointsS() As Range, rDBPointsT() As Range, rDBCertTestTemp(6) As Range, rDBPointNum(40) As Range
    Dim wireLot(6) As String, wireLotLetter(6) As String, wireLotNames(6) As String, wireLotNumber(6) As String
    Dim charNumber As Long, charNumberB As Long, charNumberA As Long, d As Long, iDBNumPoints As Long, iDBCertTemps(6) As Long
    Dim c As Long, intWireLotAmt As Long, intDistWireLotAmt As Long, i As Long, x As Long, intDistNumFileNames As Long, intDistNumUsedWireLots As Long
    Dim k As Long
    Dim bWireLotMatch As Boolean
    Dim v As Variant, v2 As Variant
    Dim obj As Object, obj2 As Object
    Dim Source_Workbook As Workbook, Source_Worksheet As Worksheet, Target_Worksheet As Worksheet

    'Assign the Workbook Path
    Dim strPath As String: strPath = GetRootDrive(ThisWorkbook.path) & "\Wires_Daqbook\"
    
    'Set Target Workbook
    Set Target_Worksheet = ThisWorkbook.Worksheets("Standards_Import")
    
    '=======================================================================================
    'Start Daqbook Data ====================================================================

    'Initialize the DaqBook information from Main Page
    Set rDBName = Worksheets("Main").Range("D9")
    Set rDBDate = Worksheets("Main").Range("D14")
    
    Select Case rDBName
        Case "J1":  iDBNumPoints = Worksheets("Standards_Info").Range("E2")
        Case "J2":  iDBNumPoints = Worksheets("Standards_Info").Range("E3")
        Case "J3":  iDBNumPoints = Worksheets("Standards_Info").Range("E4")
        Case "K2":  iDBNumPoints = Worksheets("Standards_Info").Range("E5")
        Case "K3":  iDBNumPoints = Worksheets("Standards_Info").Range("E6")
        Case "K4":  iDBNumPoints = Worksheets("Standards_Info").Range("E7")
        Case "K5":  iDBNumPoints = Worksheets("Standards_Info").Range("E8")
        Case "K6":  iDBNumPoints = Worksheets("Standards_Info").Range("E9")
        Case "N1":  iDBNumPoints = Worksheets("Standards_Info").Range("E10")
        Case "N2":  iDBNumPoints = Worksheets("Standards_Info").Range("E11")
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
    intWireLotAmt = 0

    'Check and see how many wirelots need to be processed
    For c = 0 To 5
        With Worksheets("Main").Cells(55, 4 + c)
            If .Value <> "" Then
                intWireLotAmt = intWireLotAmt + 1
                wireLotNames(c) = Trim(.Value)
            End If
        End With
    Next c

    'Remove duplicate lot numbers and create array for getting files
    Set obj = CreateObject("Scripting.Dictionary")

    For i = LBound(wireLotNames) To UBound(wireLotNames)
        obj(wireLotNames(i)) = 1
    Next i

    x = 0
    For Each v In obj.Keys()
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

    For Each v2 In obj2.Keys()
       strFileDeDupped(x) = v2
       x = x + 1
    Next v2

    'Get number of distinct file names
    intDistNumFileNames = x - 1

    'Clear all cells before we import information
    Target_Worksheet.Columns("A").ClearContents
    Target_Worksheet.Columns("C:M").ClearContents

    'Open Thermocouple file(s) and grab information
    d = 0
    For c = 0 To intDistNumFileNames - 1
        If Not IsNull(strFileDeDupped(c)) Then
    
            Set Source_Workbook = Workbooks.Open(strFileDeDupped(c))
            Set Source_Worksheet = Source_Workbook.Worksheets("TC Form")
    
            ' Check to see if this is the first run
            If c = 0 Then
                Target_Worksheet.Range("C1:G1").Value = Source_Worksheet.Range("D650:H650").Value
                Target_Worksheet.Range("H1:L1").Value = Source_Worksheet.Range("D656:H656").Value
            Else
                d = c + CLng(bWireLotMatch)
            End If
    
            ' Update Target File
            If IsInArray(Source_Worksheet.Range("B651").Value, wireLot) Then
                Target_Worksheet.Range("A" & 3 + d).Value = Source_Worksheet.Range("B651").Value
                Target_Worksheet.Range("C" & 3 + d & ":G" & 3 + d).Value = Source_Worksheet.Range("K653:O653").Value
                Target_Worksheet.Range("H" & 3 + d & ":L" & 3 + d).Value = Source_Worksheet.Range("K660:O660").Value
            Else
                d = d - 1
            End If
    
            If IsInArray(Source_Worksheet.Range("B691").Value, wireLot) And Source_Worksheet.Range("B691").Value <> 0 Then
                Target_Worksheet.Range("A" & 4 + d).Value = Source_Worksheet.Range("B691").Value
                Target_Worksheet.Range("C" & 4 + d & ":G" & 4 + d).Value = Source_Worksheet.Range("K693:O693").Value
                Target_Worksheet.Range("H" & 4 + d & ":L" & 4 + d).Value = Source_Worksheet.Range("K700:O700").Value
            Else
                bWireLotMatch = False
            End If
    
            Source_Workbook.Close False
        End If
    Next c


    'END WireLot Data ======================================================================
    '=======================================================================================
    
    'Call the Wire Correction sub
    Call Sheet14.Write_Wire_Correction_Factors

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

HandleError:
    MsgBox Err.Description & Chr(13) & "Error Number: " & Err.Number, vbExclamation
    Resume CleanExit
End Sub

Function GetUniqueValues(ByRef inputArray() As String) As Collection
    Dim dict As Object
    Dim i As Long
    Set dict = CreateObject("Scripting.Dictionary")

    For i = LBound(inputArray) To UBound(inputArray)
        If Trim(inputArray(i)) <> "" Then
            dict(Trim(inputArray(i))) = 1
        End If
    Next i

    Set GetUniqueValues = New Collection
    Dim key As Variant
    For Each key In dict.Keys
        GetUniqueValues.Add key
    Next key
End Function

