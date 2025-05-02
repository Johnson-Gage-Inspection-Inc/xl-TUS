VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Read_External_Workbook()
    'Check for errors
    ' On Error GoTo HandleError
    
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
    Dim Source_Workbook2 As Workbook
    Dim Source_Worksheet1 As Worksheet, Source_Worksheet2 As Worksheet, Standards_Import As Worksheet

    'Assign the Workbook Path
    Dim strPath As String: strPath = GetRootDrive(ThisWorkbook.path) & "\Wires_Daqbook\"
    
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
    End If
    
    'Clear all cells before we import information
    Me.Columns("O:U").ClearContents
    
    'Open Daqbook File
    Set Source_Worksheet1 = ThisWorkbook.Worksheets("DaqbookCertData")
    
    'Get Test Temps and process CF from Source
    For i = 0 To 5
        'Label the spreadsheet with the correct temp and also assign the temp to the correct processing value
        iDBCertTemps(i) = Source_Worksheet1.Range("A" & 43 + i).Value
        Set rDBCertTestTemp(i) = Me.Cells(1, 16 + i)
        rDBCertTestTemp(i) = iDBCertTemps(i)
    
        'Process each channel per given test temp
        For c = 0 To (iDBNumPoints - 1)
        
            'Label the Points
            Set rDBPointNum(c) = Me.Cells(2 + c, 15)
            rDBPointNum(c) = c + 1
            
            'Get the values
            If c < 6 Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet1.Cells(43 + i, 2 + x)
                    Set rDBPointsT(c) = Me.Cells(2 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
    
            If c > 5 And c < 12 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet1.Cells(51 + i, 2 + x)
                    Set rDBPointsT(c) = Me.Cells(8 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 11 And c < 18 And c < iDBNumPoints Then
                'If number of points is 14
                If iDBNumPoints = 14 Then
                    For x = 0 To 1
                        Set rDBPointsS(c) = Source_Worksheet1.Cells(61 + i, 2 + x)
                        Set rDBPointsT(c) = Me.Cells(14 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                Else
                    'If number of points in not 14
                    For x = 0 To 5
                        Set rDBPointsS(c) = Source_Worksheet1.Cells(61 + i, 2 + x)
                        Set rDBPointsT(c) = Me.Cells(14 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                End If
            End If
    
            If c > 17 And c < 24 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet1.Cells(69 + i, 2 + x)
                    Set rDBPointsT(c) = Me.Cells(20 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 23 And c < 30 And c < iDBNumPoints Then
                If iDBNumPoints = 28 Then
                    'If number of points is 28
                    For x = 0 To 3
                        Set rDBPointsS(c) = Source_Worksheet1.Cells(79 + i, 2 + x)
                        Set rDBPointsT(c) = Me.Cells(26 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                Else
                    'If number of points is not 28
                    For x = 0 To 5
                        Set rDBPointsS(c) = Source_Worksheet1.Cells(79 + i, 2 + x)
                        Set rDBPointsT(c) = Me.Cells(26 + x, 16 + i)
                        rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                    Next
                End If
            End If
            
            If c > 29 And c < 36 And c < iDBNumPoints Then
                For x = 0 To 5
                    Set rDBPointsS(c) = Source_Worksheet1.Cells(87 + i, 2 + x)
                    Set rDBPointsT(c) = Me.Cells(32 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
            
            If c > 35 And c < iDBNumPoints Then
                For x = 0 To 3
                    Set rDBPointsS(c) = Source_Worksheet1.Cells(97 + i, 2 + x)
                    Set rDBPointsT(c) = Me.Cells(38 + x, 16 + i)
                    rDBPointsT(c).Value = WorksheetFunction.Round((rDBPointsS(c).Value - iDBCertTemps(i)) * -1, 2)
                Next
            End If
        Next
    Next
    
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
    Me.Columns("A").ClearContents
    Me.Columns("C:M").ClearContents

    'Open Thermocouple file(s) and grab information
    d = 0
    For c = 0 To intDistNumFileNames - 1
        If Not IsNull(strFileDeDupped(c)) Then
    
            Set Source_Workbook2 = Workbooks.Open(strFileDeDupped(c))
            Set Source_Worksheet2 = Source_Workbook2.Worksheets("TC Form")
    
            ' Check to see if this is the first run
            If c = 0 Then
                Me.Range("C1:G1").Value = Source_Worksheet2.Range("D650:H650").Value
                Me.Range("H1:L1").Value = Source_Worksheet2.Range("D656:H656").Value
            Else
                d = c + CLng(bWireLotMatch)
            End If
    
            ' Update Target File
            If IsInArray(Source_Worksheet2.Range("B651").Value, wireLot) Then
                Me.Range("A" & 3 + d).Value = Source_Worksheet2.Range("B651").Value
                Me.Range("C" & 3 + d & ":G" & 3 + d).Value = Source_Worksheet2.Range("K653:O653").Value
                Me.Range("H" & 3 + d & ":L" & 3 + d).Value = Source_Worksheet2.Range("K660:O660").Value
            Else
                d = d - 1
            End If
    
            If IsInArray(Source_Worksheet2.Range("B691").Value, wireLot) And Source_Worksheet2.Range("B691").Value <> 0 Then
                Me.Range("A" & 4 + d).Value = Source_Worksheet2.Range("B691").Value
                Me.Range("C" & 4 + d & ":G" & 4 + d).Value = Source_Worksheet2.Range("K693:O693").Value
                Me.Range("H" & 4 + d & ":L" & 4 + d).Value = Source_Worksheet2.Range("K700:O700").Value
            Else
                bWireLotMatch = False
            End If
    
            Source_Workbook2.Close False
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
