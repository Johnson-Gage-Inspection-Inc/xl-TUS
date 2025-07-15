Attribute VB_Name = "TestModule1"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

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

    InputMainSheetData
End Sub
Private Sub LoadTestChannelBlock(tableName As String, startCell As String, channelCount As Long, startChannel As Long, tsvPath As String)
    With wsMain
        .Range("D26:D28").value = "9:04:00 AM"
        .Range("D30").value = "9:40:00 AM"
        .Range("D32").value = "30"
    End With
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
    
    With wsMain
        .Range("D9").value = "J2"
        .Range("K14").value = "69975"
        Sleep 1
    
        .Range("K15").value = "SIM Load Hot"
        .Range("D48").value = "J01-J24"
        .Range("D51").value = "10"
        .Range("D52").value = "0"
        .Range("D56").value = "10"
        .Range("D57").value = ""
        .Range("D3").value = "2/17/2025"
        .Range("D17:D18").value = "10"
        .Range("D22").value = "68"
        .Range("D23").value = "19"
        .Range("D24").value = "1"
        .Range("D15:D16").value = "100"

        Dim i As Long
        For i = 1 To 10
            .Range("O" & (i + 4)).value = "J" & Format(i, "00")
        Next i
    End With

End Sub
Private Sub PopulateComparisonReportInputs()
    ' Set up Comparison Report data in B37:G40
    With wsMain
        .Range("B37").value = 10
        .Range("C37").value = "Controller"
        .Range("D37").value = 102
        .Range("E37").value = 102
        .Range("F37").value = 103
        .Range("G37").value = 102
    
        .Range("B38").value = 10
        .Range("C38").value = "Recorder"
        .Range("D38").value = 102.44
        .Range("E38").value = 103.45
        .Range("F38").value = 104.13
        .Range("G38").value = 103.45
    End With
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
Private Sub TCAlerts_ContainsExpectedHighLowOnly():
    Application.ScreenUpdating = True  ' For testing purposes
    Application.EnableEvents = True  ' For testing purposes
    Sleep 1
    
    On Error GoTo TestFail  ' Press F9 to add a breakpoint here, to test with data
    ' Then, Press F5 to clear the data after testing.

    PopulateComparisonReportInputs
    ' Inject raw test data from file
    LoadTestChannelBlock "DataForChannels1to14", "A2", 14, 1, "C:\Users\JeffHall\git\xl-TUS\test1.tsv"
    
    Sleep 1

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


