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


