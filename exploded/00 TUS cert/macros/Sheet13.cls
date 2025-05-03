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

    ThisWorkbook.Connections("Query - DaqbookOffsets").Refresh

    'Check for errors
    On Error GoTo HandleError
    
    'Stop Screen Updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
     
    'Define Object for Target Workbook
    Dim strFile(6) As String, strFileDeDupped(6) As String
    Dim strFileHold1 As String, strFileHold2 As String, strFileHold3 As String
    Dim wireLot(6) As String, wireLotLetter(6) As String, wireLotNames(6) As String, wireLotNumber(6) As String
    Dim charNumber As Long, charNumberB As Long, charNumberA As Long
    Dim c As Long, i As Long, x As Long, d As Long
    Dim intWireLotAmt As Long, intDistWireLotAmt As Long, intDistNumFileNames As Long, intDistNumUsedWireLots As Long
    Dim bWireLotMatch As Boolean
    Dim v As Variant, v2 As Variant
    Dim obj As Object, obj2 As Object
    Dim Source_Workbook2 As Workbook
    Dim Source_Worksheet2 As Worksheet

    'Assign the Workbook Path
    Dim strPath As String: strPath = GetRootDrive(ThisWorkbook.path) & "\Wires_Daqbook\"
                
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
