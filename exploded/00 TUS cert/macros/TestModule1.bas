Attribute VB_Name = "TestModule1"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub IsViewOnlySheet_ReturnsTrue_ForSheet21()
    Dim ws As Worksheet
    Set ws = Sheet21 ' Or: Set ws = ThisWorkbook.Sheets("Customers") if renamed

    Assert.IsTrue IsViewOnlySheet(ws), "Expected Sheet21 to be view-only"
End Sub

'@TestMethod
Public Sub QRound_RoundsUpCorrectly()
    ' Test rounding up
    Dim result As Double
    result = QRound(1.15)
    Assert.AreEqual 1.2, result, "Expected 1.15 to round up to 1.2"
End Sub

'@TestMethod
Public Sub QRound_RoundsDownCorrectly()
    ' Test rounding down
    Dim result As Double
    result = QRound(1.14)
    Assert.AreEqual 1.1, result, "Expected 1.14 to round down to 1.1"
End Sub

'@TestMethod
Public Sub QRound_RoundsToEvenForTies()
    ' Test rounding to the nearest even number for ties
    Dim result As Double
    result = QRound(1.25)
    Assert.AreEqual 1.2, result, "Expected 1.25 to round to the nearest even number 1.2"
    
    result = QRound(1.35)
    Assert.AreEqual 1.4, result, "Expected 1.35 to round to the nearest even number 1.4"
End Sub

'@TestMethod
Public Sub QRound_RoundsNegativeNumbersCorrectly()
    ' Test rounding for negative numbers
    Dim result As Double
    result = QRound(-1.15)
    Assert.AreEqual -1.2, result, "Expected -1.15 to round up to -1.2"
    
    result = QRound(-1.25)
    Assert.AreEqual -1.2, result, "Expected -1.25 to round to the nearest even number -1.2"
End Sub

' Helper function to clear all test input data
Private Sub ClearTestInputs()
    Dim wsMain As Worksheet
    Dim wsDaqBook As Worksheet
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set wsDaqBook = ThisWorkbook.Sheets("DaqBook_RAW_Data")
    
    ' Clear inputs in Main sheet
    wsMain.Range("D3").ClearContents
    wsMain.Range("D7").ClearContents
    wsMain.Range("D9").ClearContents
    wsMain.Range("D15:D19").ClearContents ' Consecutive cells
    wsMain.Range("D22").ClearContents
    wsMain.Range("D23").ClearContents
    wsMain.Range("D26:D28").ClearContents ' Consecutive cells
    wsMain.Range("D30").ClearContents
    wsMain.Range("D32").ClearContents
    wsMain.Range("K14:L14").ClearContents ' Merged cells
    wsMain.Range("K15:L15").ClearContents ' Merged cells
    wsMain.Range("D48").ClearContents
    wsMain.Range("D51").ClearContents
    wsMain.Range("D52").ClearContents
    wsMain.Range("D56").ClearContents
    wsMain.Range("D57").ClearContents
    wsMain.Range("O5:O14").ClearContents ' Consecutive cells

    ' Clear inputs in DaqBook_RAW_Data sheet
    wsDaqBook.Range("A2:K38").ClearContents
End Sub

'@TestMethod("TUS Input and Clear Test")
Private Sub TUS_InputAndClearTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim wsMain As Worksheet
    Dim wsDaqBook As Worksheet
    Set wsMain = ThisWorkbook.Sheets("Main")
    Set wsDaqBook = ThisWorkbook.Sheets("DaqBook_RAW_Data")
    
    'Act:
    ' Input data into Main sheet
    wsMain.Range("D3").Value = "2/17/2025"
    wsMain.Range("D7").Value = "Danny Turkali"
    wsMain.Range("D9").Value = "J2"
    wsMain.Range("D15").Value = "100"
    wsMain.Range("D16").Value = "100"
    wsMain.Range("D17").Value = "10"
    wsMain.Range("D18").Value = "10"
    wsMain.Range("D19").Value = "10"
    wsMain.Range("D22").Value = "68"
    wsMain.Range("D23").Value = "19"
    wsMain.Range("D26").Value = "9:04:00 AM"
    wsMain.Range("D27").Value = "9:04:00 AM"
    wsMain.Range("D28").Value = "9:04:00 AM"
    wsMain.Range("D30").Value = "9:40:00 AM"
    wsMain.Range("D32").Value = "30"
    wsMain.Range("K14").Value = "56561-069975-01"
    wsMain.Range("K15").Value = "SIM Load Hot"
    wsMain.Range("D48").Value = "J01-J24"
    wsMain.Range("D51").Value = "10"
    wsMain.Range("D52").Value = "0"
    wsMain.Range("D56").Value = "10"
    wsMain.Range("D57").Value = ""
    wsMain.Range("O5").Value = "J01"
    wsMain.Range("O6").Value = "J02"
    wsMain.Range("O7").Value = "J03"
    wsMain.Range("O8").Value = "J04"
    wsMain.Range("O9").Value = "J05"
    wsMain.Range("O10").Value = "J06"
    wsMain.Range("O11").Value = "J07"
    wsMain.Range("O12").Value = "J08"
    wsMain.Range("O13").Value = "J09"
    wsMain.Range("O14").Value = "J10"

    ' Input data into DaqBook_RAW_Data sheet
    Dim rowIdx As Long, colIdx As Long
    Dim data As Variant
    data = Split(CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\JeffHall\git\xl-TUS\test1.tsv").ReadAll, vbCrLf)
    For rowIdx = LBound(data) To UBound(data) - 1
        If Len(Trim(data(rowIdx))) > 0 Then
            Dim values As Variant
            values = Split(data(rowIdx), vbTab)
            For colIdx = LBound(values) To UBound(values)
                wsDaqBook.Cells(rowIdx + 2, colIdx + 1).Value = values(colIdx)
            Next colIdx
        End If
    Next rowIdx    ' Clear inputs using helper function
    ClearTestInputs

    'Assert:
    Assert.IsTrue wsMain.Range("D3").Value = "", "Expected D3 to be cleared"
    Assert.IsTrue wsDaqBook.Range("A2").Value = "", "Expected A2 to be cleared"
    Assert.Succeed

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

