Attribute VB_Name = "TestModule1"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private wsMain As Worksheet
Private wsDaqBook As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")

    Set wsMain = ThisWorkbook.Sheets("Main")
    Set wsDaqBook = ThisWorkbook.Sheets("DaqBook_RAW_Data")

    Application.ScreenUpdating = False
    InputMainSheetData

    ' Inject raw test data from file
    LoadTestChannelBlock "DataForChannels1to14", "A2", 14, 1, "C:\Users\JeffHall\git\xl-TUS\test1.tsv"
    
    Application.ScreenUpdating = True
End Sub
Private Sub LoadTestChannelBlock(tableName As String, startCell As String, channelCount As Long, startChannel As Long, tsvPath As String)
    Dim rawText As String
    rawText = CreateObject("Scripting.FileSystemObject").OpenTextFile(tsvPath).ReadAll
    Sheet7.PasteChannelBlock "DaqBook_RAW_Data", startCell, channelCount, tableName, startChannel, rawText
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    ClearMainSheetInputs
    Sheet7.TruncateChannels1to14
    Sheet7.TruncateChannels15to28
    Sheet7.TruncateChannels29to40

    Set wsDaqBook = Nothing
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

Private Sub InputMainSheetData()
    wsMain.Range("D3").Value = "2/17/2025"
    wsMain.Range("D9").Value = "J2"
    wsMain.Range("D15:D16").Value = "100"
    wsMain.Range("D17:D19").Value = "10"
    wsMain.Range("D22").Value = "68"
    wsMain.Range("D23").Value = "19"
    wsMain.Range("D24").Value = "1"
    wsMain.Range("D26:D28").Value = "9:04:00 AM"
    wsMain.Range("D30").Value = "9:40:00 AM"
    wsMain.Range("D32").Value = "30"
    wsMain.Range("K14").Value = "56561-069975"
    wsMain.Range("K15").Value = "SIM Load Hot"
    wsMain.Range("D48").Value = "J01-J24"
    wsMain.Range("D51").Value = "10"
    wsMain.Range("D52").Value = "0"
    wsMain.Range("D56").Value = "10"
    wsMain.Range("D57").Value = ""

    Dim i As Long
    For i = 1 To 10
        wsMain.Range("O" & (i + 4)).Value = "J" & Format(i, "00")
    Next i

    ' Set up temperature test data in B37:G40
    wsMain.Range("B37").Value = 10
    wsMain.Range("C37").Value = "Controller"
    wsMain.Range("D37").Value = 374
    wsMain.Range("E37").Value = 375
    wsMain.Range("F37").Value = 375
    wsMain.Range("G37").Value = 375

    wsMain.Range("B38").Value = 10
    wsMain.Range("C38").Value = "Recorder 1"
    wsMain.Range("D38").Value = 375.8
    wsMain.Range("E38").Value = 375.8
    wsMain.Range("F38").Value = 375.8
    wsMain.Range("G38").Value = 375.8

    wsMain.Range("B39").Value = 7
    wsMain.Range("C39").Value = "Rec 2 High"
    wsMain.Range("D39").Value = 372.8
    wsMain.Range("E39").Value = 372.7
    wsMain.Range("F39").Value = 372.9
    wsMain.Range("G39").Value = 372.8

    wsMain.Range("B40").Value = 5
    wsMain.Range("C40").Value = "Rec 3 Low"
    wsMain.Range("D40").Value = 372.5
    wsMain.Range("E40").Value = 372.6
    wsMain.Range("F40").Value = 372.8
    wsMain.Range("G40").Value = 372.8
End Sub


Private Sub ClearMainSheetInputs()
    With wsMain
        .Range("D3,D9,D22,D23,D30,D32,D48,D51,D52,D56,D57").ClearContents
        .Range("D15:D19").ClearContents
        .Range("D26:D28").ClearContents
        .Range("K14:L14").ClearContents
        .Range("K15:L15").ClearContents
        .Range("O5:O14").ClearContents
        .Range("B37:L44").ClearContents
    End With
End Sub

' Reusable test for PasteChannelsXXtoYY
Private Sub TestPasteRoutine(pasteSubName As String, expectedStartCol As String, expectedFirstChannel As Long)
    Dim testTSVPath As String
    testTSVPath = "C:\Users\JeffHall\git\xl-TUS\test1.tsv" ' Or parameterize further
    
    ' Load test content into clipboard (PasteChannels still expects clipboard use)

    ' Ensure clipboard is populated
    Dim rawText As String
    rawText = CreateObject("Scripting.FileSystemObject").OpenTextFile(testTSVPath).ReadAll
    With CreateObject("MSForms.DataObject")
        .SetText rawText
        .PutInClipboard
    End With

    ' Dynamically invoke the macro by name
    Application.Run pasteSubName

    ' Verify paste occurred
    Dim pasteCell As Range
    Set pasteCell = wsDaqBook.Range(expectedStartCol).Offset(1, 0)

    Assert.IsTrue IsDate(pasteCell.Value), "Expected time value in " & pasteCell.Address
    Assert.IsTrue IsNumeric(pasteCell.Offset(0, 1).Value), "Expected numeric channel in " & pasteCell.Offset(0, 1).Address
    Assert.AreEqual expectedFirstChannel, CLng(wsDaqBook.Range(expectedStartCol).Offset(0, 1).Value), "Expected first channel label"

    ' Add further range checks here
End Sub

'@TestMethod("Main Sheet Logic")
Private Sub TCAlerts_ContainsExpectedHighLowOnly()
    On Error GoTo TestFail

    Dim i As Long
    Dim val As Variant

    For i = 5 To 14
        val = wsMain.Range("P" & i).Value

        ' Ensure no "Dropped"
        Assert.AreNotEqual "Dropped", val, "Expected P" & i & " not to contain 'Dropped'"

        ' Check expected values
        Select Case i
            Case 6
                Assert.AreEqual "High", val, "Expected P6 to be High"
            Case 8
                Assert.AreEqual "Low", val, "Expected P8 to be Low"
            Case Else
                Assert.AreEqual "", val, "Expected P" & i & " to be empty"
        End Select
    Next i

    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

