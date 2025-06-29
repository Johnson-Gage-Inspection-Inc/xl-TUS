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
    wsMain.Range("D3").value = "2/17/2025"
    wsMain.Range("D9").value = "J2"
    wsMain.Range("D15:D16").value = "100"
    wsMain.Range("D17:D18").value = "10"
    wsMain.Range("D22").value = "68"
    wsMain.Range("D23").value = "19"
    wsMain.Range("D24").value = "1"
    wsMain.Range("D26:D28").value = "9:04:00 AM"
    wsMain.Range("D30").value = "9:40:00 AM"
    wsMain.Range("D32").value = "30"
    wsMain.Range("K14").value = "56561-069975"
    wsMain.Range("K15").value = "SIM Load Hot"
    wsMain.Range("D48").value = "J01-J24"
    wsMain.Range("D51").value = "10"
    wsMain.Range("D52").value = "0"
    wsMain.Range("D56").value = "10"
    wsMain.Range("D57").value = ""

    Dim i As Long
    For i = 1 To 10
        wsMain.Range("O" & (i + 4)).value = "J" & Format(i, "00")
    Next i

    ' Set up temperature test data in B37:G40
    wsMain.Range("B37").value = 10
    wsMain.Range("C37").value = "Controller"
    wsMain.Range("D37").value = 374
    wsMain.Range("E37").value = 375
    wsMain.Range("F37").value = 375
    wsMain.Range("G37").value = 375

    wsMain.Range("B38").value = 10
    wsMain.Range("C38").value = "Recorder 1"
    wsMain.Range("D38").value = 375.8
    wsMain.Range("E38").value = 375.8
    wsMain.Range("F38").value = 375.8
    wsMain.Range("G38").value = 375.8

    wsMain.Range("B39").value = 7
    wsMain.Range("C39").value = "Rec 2 High"
    wsMain.Range("D39").value = 372.8
    wsMain.Range("E39").value = 372.7
    wsMain.Range("F39").value = 372.9
    wsMain.Range("G39").value = 372.8

    wsMain.Range("B40").value = 5
    wsMain.Range("C40").value = "Rec 3 Low"
    wsMain.Range("D40").value = 372.5
    wsMain.Range("E40").value = 372.6
    wsMain.Range("F40").value = 372.8
    wsMain.Range("G40").value = 372.8
End Sub


Private Sub ClearMainSheetInputs()
    With wsMain
        .Range("D3,D9,D22,D23,D30,D32,D48,D51,D52,D56,D57").ClearContents
        .Range("D15:D18").ClearContents
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

    Assert.IsTrue IsDate(pasteCell.value), "Expected time value in " & pasteCell.Address
    Assert.IsTrue IsNumeric(pasteCell.Offset(0, 1).value), "Expected numeric channel in " & pasteCell.Offset(0, 1).Address
    Assert.AreEqual expectedFirstChannel, CLng(wsDaqBook.Range(expectedStartCol).Offset(0, 1).value), "Expected first channel label"

    ' Add further range checks here
End Sub

'@TestMethod("Main Sheet Logic")
Private Sub TCAlerts_ContainsExpectedHighLowOnly()
    On Error GoTo TestFail

    Dim i As Long
    Dim val As Variant

    For i = 5 To 14
        val = wsMain.Range("P" & i).value

        ' Ensure no "Dropped"
        Assert.AreNotEqual "Dropped", val, "Expected P" & i & " not to contain 'Dropped'"

        ' Check expected values
        Select Case i
            Case 6
                Assert.AreEqual "Low", val, "Expected P6 to be High"
            Case 14
                Assert.AreEqual "High", val, "Expected P8 to be Low"
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

