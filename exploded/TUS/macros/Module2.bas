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
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.Row Step -1
        If Cells(k, TestCol).value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k
    
    Worksheets("Data_Sheet_15_28").Activate
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.Row Step -1
        If Cells(k, TestCol).value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k

    Worksheets("Data_Sheet_29_40").Activate
    For k = ActiveSheet.UsedRange.Rows.count To ActiveSheet.UsedRange.Row Step -1
        If Cells(k, TestCol).value <> "" Then Exit For
    Next k
    ActiveSheet.PageSetup.PrintArea = "A1:P" & k
    
    Worksheets("Main").Activate
    
    'Set each worksheet for printing
    For Each rng In Sheets("Standards_Info").Range("U2:U9")
        If Trim(rng.value) <> "" Then
            On Error Resume Next
            Set wks = Nothing
            Set wks = Sheets(rng.value)
            On Error GoTo 0
            If wks Is Nothing Then
                MsgBox "Sheet " & rng.value & " does not exist"
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
