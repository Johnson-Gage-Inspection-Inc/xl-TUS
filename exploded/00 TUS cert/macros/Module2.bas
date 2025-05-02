Attribute VB_Name = "Module2"

Sub Printselection()
    
    'Check for errors
    On Error Resume Next
    
    'Stop Screen Updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
        
    'Assign Variables
    Dim rng As Range
    Dim wks As Worksheet
    Dim arr() As String
    Dim printSheets As Variant
    Dim i As Long: i = 0
    Dim k As Long
    Dim te As Long
    Dim TestCol As Long
        
    'Set print area for each of the Data sheets
    TestCol = 2   '<<< 1=A, 2=B, 3=C, etc
    
    Worksheets("Data_Sheet").Activate
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.row Step -1
        If Cells(k, TestCol).Value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k
    
    Worksheets("Data_Sheet_15_28").Activate
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.row Step -1
        If Cells(k, TestCol).Value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k

    Worksheets("Data_Sheet_29_40").Activate
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.row Step -1
        If Cells(k, TestCol).Value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k
    
    Worksheets("Main").Activate
    
    'Set each worksheet for printing
    For Each rng In Sheets("Standards_Info").Range("U2:U9")
        If Trim(rng.Value) <> "" Then
            On Error Resume Next
            Set wks = Nothing
            Set wks = Sheets(rng.Value)
            On Error GoTo 0
            If wks Is Nothing Then
                MsgBox "Sheet " & rng.Value & " does not exist"
            Else
                ReDim Preserve arr(i)
                arr(i) = wks.Name
                i = i + 1
            End If
        End If
    Next rng

    printSheets = arr
    Worksheets(printSheets).PrintOut Preview:=False, ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, PrToFileName:=PSFileName, IgnorePrintAreas:=False
    
    'Start Screen Updates
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

'~~> Function required to find the list from Col B
Function FindRange(FirstRange As Range, StrSearch As String, iOffset As Long) As String

    Dim aCell As Range, bCell As Range, oRange As Range, rListStart As Range, rHold As Range
    Dim ExitLoop As Boolean
    Dim strTemp As String
    Dim i As Long
    
    'Find the first cell with the same company anme
    Set aCell = FirstRange.Find(what:=StrSearch, LookIn:=xlValues, _
    Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)
    
    'Set the start variables
    i = 0
    ExitLoop = False
 
    If Not aCell Is Nothing Then
    
        'Set the start cell
        Set rListStart = aCell.Offset(, iOffset)
    
        'Set the temp value to the current cell
        Set bCell = aCell
        
        'Loop through the results
        Do While ExitLoop = False
            
            'Continue the find
            Set aCell = FirstRange.FindNext(After:=aCell)
 
            'Check to see if there is another cell with the same company name
            If Not aCell Is Nothing Then
                
                If aCell.Address = bCell.Address Then Exit Do
                i = i + 1
                
            Else
                ExitLoop = True
            End If
        Loop
        
        'Return the result
        FindRange = rListStart.Address & ":" & rListStart.Offset(i).Address
        
    End If
End Function





